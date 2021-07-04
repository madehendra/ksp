VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form MstRegisterNasabah 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MASTER REGISTER NASABAH"
   ClientHeight    =   6225
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11805
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6225
   ScaleWidth      =   11805
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   5595
      Left            =   0
      Top             =   0
      Width           =   11805
      _ExtentX        =   20823
      _ExtentY        =   9869
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
      Begin TabDlg.SSTab SSTab1 
         Height          =   5460
         Left            =   75
         TabIndex        =   0
         Top             =   45
         Width           =   11685
         _ExtentX        =   20611
         _ExtentY        =   9631
         _Version        =   393216
         Style           =   1
         Tabs            =   2
         TabHeight       =   520
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "DATA SOSIAL NASABAH"
         TabPicture(0)   =   "MstRegisterNasabah.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "BiSAFrame2"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "ALAMAT RUMAH / KANTOR"
         TabPicture(1)   =   "MstRegisterNasabah.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "BiSAFrame8"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "BiSAFrame6"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "BiSAFrame5"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).Control(3)=   "BiSAFrame4"
         Tab(1).Control(3).Enabled=   0   'False
         Tab(1).ControlCount=   4
         Begin BiSAFramProject.BiSAFrame BiSAFrame8 
            Height          =   765
            Left            =   -74895
            Top             =   4560
            Width           =   5955
            _ExtentX        =   10504
            _ExtentY        =   1349
            Caption         =   "NPWP"
            BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   -2147483633
            Begin BiSATextBoxProject.BiSATextBox cNPWP 
               Height          =   330
               Left            =   120
               TabIndex        =   1
               Top             =   255
               Width           =   4230
               _ExtentX        =   7461
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
               Caption         =   "NPWP"
               CaptionWidth    =   1300
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
         Begin BiSAFramProject.BiSAFrame BiSAFrame6 
            Height          =   2055
            Left            =   -74895
            Top             =   2400
            Width           =   5955
            _ExtentX        =   10504
            _ExtentY        =   3625
            Caption         =   "ALAMAT KANTOR"
            BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   -2147483633
            Begin BiSATextBoxProject.BiSATextBox cAlamatKantor 
               Height          =   330
               Left            =   75
               TabIndex        =   2
               Top             =   360
               Width           =   5610
               _ExtentX        =   9895
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
               Caption         =   "ALAMAT"
               CaptionWidth    =   1300
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
            Begin BiSATextBoxProject.BiSATextBox cTeleponKantor 
               Height          =   330
               Left            =   75
               TabIndex        =   3
               Top             =   720
               Width           =   3870
               _ExtentX        =   6826
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
               Caption         =   "TELEPON"
               CaptionWidth    =   1300
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
            Begin BiSATextBoxProject.BiSATextBox cFaxKantor 
               Height          =   330
               Left            =   75
               TabIndex        =   4
               Top             =   1080
               Width           =   3870
               _ExtentX        =   6826
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
               Caption         =   "FAXIMILE"
               CaptionWidth    =   1300
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
            Begin BiSATextBoxProject.BiSATextBox cKodePosKantor 
               Height          =   330
               Left            =   75
               TabIndex        =   5
               Top             =   1440
               Width           =   2340
               _ExtentX        =   4128
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
               Caption         =   "KODE POS"
               CaptionWidth    =   1300
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
         Begin BiSAFramProject.BiSAFrame BiSAFrame5 
            Height          =   1800
            Left            =   -74895
            Top             =   525
            Width           =   5955
            _ExtentX        =   10504
            _ExtentY        =   3175
            Caption         =   "ALAMAT TINGGAL"
            BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   -2147483633
            Begin BiSATextBoxProject.BiSABrowse cWilayah 
               Height          =   330
               Left            =   105
               TabIndex        =   6
               Top             =   1110
               Width           =   2460
               _ExtentX        =   4339
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
               Caption         =   "WILAYAH"
               CaptionWidth    =   1300
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
            Begin BiSATextBoxProject.BiSATextBox cAlamatRumah 
               Height          =   330
               Left            =   105
               TabIndex        =   7
               Top             =   375
               Width           =   5685
               _ExtentX        =   10028
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
               Caption         =   "ALAMAT"
               CaptionWidth    =   1300
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
            Begin BiSATextBoxProject.BiSATextBox cTeleponRumah 
               Height          =   330
               Left            =   105
               TabIndex        =   8
               Top             =   735
               Width           =   4230
               _ExtentX        =   7461
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
               Caption         =   "TELEPON"
               CaptionWidth    =   1300
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
            Begin BiSATextBoxProject.BiSATextBox cNamaWilayah 
               Height          =   330
               Left            =   2580
               TabIndex        =   9
               Top             =   1110
               Width           =   3255
               _ExtentX        =   5741
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
         Begin BiSAFramProject.BiSAFrame BiSAFrame2 
            Height          =   4905
            Left            =   60
            Top             =   435
            Width           =   11490
            _ExtentX        =   20267
            _ExtentY        =   8652
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
            Begin BiSAFramProject.BiSAFrame BiSAFrame7 
               Height          =   675
               Left            =   4155
               Top             =   210
               Width           =   4230
               _ExtentX        =   7461
               _ExtentY        =   1191
               Caption         =   "Jenis Keanggotaan"
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
               Begin VB.OptionButton optJenisAnggota 
                  Caption         =   "&1 Anggota Biasa"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   300
                  Index           =   0
                  Left            =   150
                  TabIndex        =   45
                  Top             =   225
                  Width           =   1785
               End
               Begin VB.OptionButton optJenisAnggota 
                  Caption         =   "&2 Calon Anggota"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   300
                  Index           =   1
                  Left            =   1965
                  TabIndex        =   44
                  Top             =   225
                  Width           =   2145
               End
            End
            Begin VB.OptionButton optStatusPerkawinan 
               Caption         =   "&2. BELUM"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   1
               Left            =   4020
               TabIndex        =   22
               TabStop         =   0   'False
               Top             =   2880
               Width           =   1065
            End
            Begin VB.OptionButton optStatusPerkawinan 
               Caption         =   "&1. KAWIN"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   0
               Left            =   2715
               TabIndex        =   21
               TabStop         =   0   'False
               Top             =   2880
               Width           =   1125
            End
            Begin BiSAFramProject.BiSAFrame BiSAFrame11 
               Height          =   375
               Left            =   2715
               Top             =   1710
               Width           =   3180
               _ExtentX        =   5609
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
               BorderStyle     =   0
               BackColor       =   -2147483633
               Begin VB.OptionButton OptGolonganDarah 
                  Caption         =   "A"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Index           =   0
                  Left            =   90
                  TabIndex        =   13
                  TabStop         =   0   'False
                  Top             =   60
                  Width           =   570
               End
               Begin VB.OptionButton OptGolonganDarah 
                  Caption         =   "B"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Index           =   1
                  Left            =   735
                  TabIndex        =   12
                  TabStop         =   0   'False
                  Top             =   60
                  Width           =   570
               End
               Begin VB.OptionButton OptGolonganDarah 
                  Caption         =   "AB"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Index           =   2
                  Left            =   1350
                  TabIndex        =   11
                  TabStop         =   0   'False
                  Top             =   60
                  Width           =   570
               End
               Begin VB.OptionButton OptGolonganDarah 
                  Caption         =   "O"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Index           =   3
                  Left            =   2040
                  TabIndex        =   10
                  TabStop         =   0   'False
                  Top             =   60
                  Width           =   570
               End
            End
            Begin BiSAFramProject.BiSAFrame BiSAFrame10 
               Height          =   375
               Left            =   2715
               Top             =   1320
               Width           =   3180
               _ExtentX        =   5609
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
               BorderStyle     =   0
               BackColor       =   -2147483633
               Begin VB.OptionButton optSex 
                  Caption         =   "&LAKI-LAKI"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   270
                  Index           =   0
                  Left            =   90
                  TabIndex        =   15
                  TabStop         =   0   'False
                  Top             =   60
                  Width           =   1290
               End
               Begin VB.OptionButton optSex 
                  Caption         =   "&PEREMPUAN"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   270
                  Index           =   1
                  Left            =   1545
                  TabIndex        =   14
                  TabStop         =   0   'False
                  Top             =   75
                  Width           =   1455
               End
            End
            Begin BiSATextBoxProject.BiSATextBox cStatusPerkawinan 
               Height          =   300
               Left            =   5310
               TabIndex        =   16
               Top             =   2790
               Visible         =   0   'False
               Width           =   510
               _ExtentX        =   900
               _ExtentY        =   529
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
            Begin BiSATextBoxProject.BiSATextBox cKode 
               Height          =   330
               Left            =   3150
               TabIndex        =   17
               Top             =   240
               Width           =   915
               _ExtentX        =   1614
               _ExtentY        =   582
               Text            =   "123456"
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
               MaxLength       =   6
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
            Begin BiSATextBoxProject.BiSATextBox cNPekerjaan 
               Height          =   315
               Left            =   3780
               TabIndex        =   18
               Top             =   4275
               Width           =   3150
               _ExtentX        =   5556
               _ExtentY        =   556
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
            Begin BiSATextBoxProject.BiSATextBox cNAgama 
               Height          =   315
               Left            =   3780
               TabIndex        =   19
               Top             =   3900
               Width           =   3150
               _ExtentX        =   5556
               _ExtentY        =   556
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
            Begin BiSATextBoxProject.BiSABrowse cAgama 
               Height          =   330
               Left            =   600
               TabIndex        =   20
               Top             =   3900
               Width           =   3180
               _ExtentX        =   5609
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
               Caption         =   "AGAMA"
               CaptionWidth    =   2000
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
            Begin BiSATextBoxProject.BiSATextBox cGolDarah 
               Height          =   300
               Left            =   6435
               TabIndex        =   23
               Top             =   1710
               Visible         =   0   'False
               Width           =   570
               _ExtentX        =   1005
               _ExtentY        =   529
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
            Begin BiSATextBoxProject.BiSATextBox cSex 
               Height          =   300
               Left            =   5925
               TabIndex        =   24
               Top             =   1710
               Visible         =   0   'False
               Width           =   480
               _ExtentX        =   847
               _ExtentY        =   529
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
               Left            =   600
               TabIndex        =   25
               Top             =   960
               Width           =   7065
               _ExtentX        =   12462
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
               Caption         =   "NAMA LENGKAP"
               CaptionWidth    =   2000
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
            Begin BiSADateProject.BiSADate dTglRegister 
               Height          =   330
               Left            =   600
               TabIndex        =   26
               Top             =   600
               Width           =   3480
               _ExtentX        =   6138
               _ExtentY        =   582
               BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   -2147483640
               Caption         =   "TGL REGISTER"
               CaptionWidth    =   2000
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
            Begin BiSATextBoxProject.BiSATextBox cCabang 
               Height          =   330
               Left            =   600
               TabIndex        =   27
               Top             =   240
               Width           =   2535
               _ExtentX        =   4471
               _ExtentY        =   582
               Text            =   "12"
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
               MaxLength       =   2
               Caption         =   "NO. REGISTER"
               CaptionWidth    =   2000
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
            Begin BiSATextBoxProject.BiSATextBox cTempatLahir 
               Height          =   330
               Left            =   600
               TabIndex        =   28
               Top             =   2115
               Width           =   5340
               _ExtentX        =   9419
               _ExtentY        =   582
               Text            =   "12"
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
               MaxLength       =   20
               Caption         =   "TEMPAT LAHIR"
               CaptionWidth    =   2000
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
            Begin BiSADateProject.BiSADate dTglLahir 
               Height          =   330
               Left            =   600
               TabIndex        =   29
               Top             =   2490
               Width           =   3480
               _ExtentX        =   6138
               _ExtentY        =   582
               BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   -2147483640
               Caption         =   "TANGGAL LAHIR"
               CaptionWidth    =   2000
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
            Begin BiSATextBoxProject.BiSATextBox cKTP 
               Height          =   330
               Left            =   600
               TabIndex        =   30
               Top             =   3180
               Width           =   5340
               _ExtentX        =   9419
               _ExtentY        =   582
               Text            =   "12"
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
               Caption         =   "NO. SIM/KTP"
               CaptionWidth    =   2000
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
            Begin BiSADateProject.BiSADate dTglKTP 
               Height          =   330
               Left            =   600
               TabIndex        =   31
               Top             =   3540
               Width           =   3480
               _ExtentX        =   6138
               _ExtentY        =   582
               BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   -2147483640
               Caption         =   "TGL BERLAKU IDENT."
               CaptionWidth    =   2000
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
            Begin BiSATextBoxProject.BiSABrowse cPekerjaan 
               Height          =   330
               Left            =   600
               TabIndex        =   32
               Top             =   4260
               Width           =   3180
               _ExtentX        =   5609
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
               Caption         =   "PEKERJAAN"
               CaptionWidth    =   2000
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
               Caption         =   "JENIS KELAMIN"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   600
               TabIndex        =   35
               Top             =   1395
               Width           =   1650
            End
            Begin VB.Label Label2 
               Caption         =   "GOLONGAN DARAH"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   600
               TabIndex        =   34
               Top             =   1770
               Width           =   1785
            End
            Begin VB.Label Label3 
               Caption         =   "STATUS PERKAWINAN"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   600
               TabIndex        =   33
               Top             =   2865
               Width           =   2010
            End
         End
         Begin BiSAFramProject.BiSAFrame BiSAFrame4 
            Height          =   4695
            Left            =   -68895
            Top             =   615
            Width           =   5505
            _ExtentX        =   9710
            _ExtentY        =   8281
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
            Begin MSComDlg.CommonDialog CommonDialog2 
               Left            =   3225
               Top             =   1050
               _ExtentX        =   847
               _ExtentY        =   847
               _Version        =   393216
            End
            Begin MSComDlg.CommonDialog CommonDialog1 
               Left            =   930
               Top             =   930
               _ExtentX        =   847
               _ExtentY        =   847
               _Version        =   393216
               DialogTitle     =   "Foto Nasabah"
               FilterIndex     =   1
            End
            Begin BiSAButtonProject.BiSAButton cmdFoto 
               Height          =   435
               Left            =   105
               TabIndex        =   36
               Top             =   3540
               Width           =   2190
               _ExtentX        =   3863
               _ExtentY        =   767
               Caption         =   "FOTO"
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
            End
            Begin BiSAButtonProject.BiSAButton cmdTTD 
               Height          =   435
               Left            =   2325
               TabIndex        =   37
               Top             =   3540
               Width           =   3105
               _ExtentX        =   5477
               _ExtentY        =   767
               Caption         =   "SPECIMEN"
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
            End
            Begin VB.Image Image2 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   2475
               Left            =   2325
               Stretch         =   -1  'True
               Top             =   975
               Width           =   3105
            End
            Begin VB.Image Image1 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   2475
               Left            =   105
               Stretch         =   -1  'True
               Top             =   975
               Width           =   2190
            End
         End
      End
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame3 
      Height          =   630
      Left            =   0
      Top             =   5595
      Width           =   11805
      _ExtentX        =   20823
      _ExtentY        =   1111
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
      Begin BiSAButtonProject.BiSAButton cmdHapus 
         Height          =   435
         Left            =   2190
         TabIndex        =   38
         Top             =   105
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   767
         Caption         =   "    &Delete"
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
         Picture         =   "MstRegisterNasabah.frx":0038
      End
      Begin BiSAButtonProject.BiSAButton cmdAktivasi 
         Height          =   435
         Left            =   3360
         TabIndex        =   39
         Top             =   105
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
         Picture         =   "MstRegisterNasabah.frx":02C2
      End
      Begin BiSAButtonProject.BiSAButton cmdSimpan 
         Height          =   435
         Left            =   9510
         TabIndex        =   40
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
         Picture         =   "MstRegisterNasabah.frx":0461
      End
      Begin BiSAButtonProject.BiSAButton cmdEdit 
         Height          =   435
         Left            =   1140
         TabIndex        =   41
         Top             =   105
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   767
         Caption         =   "  &Edit"
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
         Picture         =   "MstRegisterNasabah.frx":0877
      End
      Begin BiSAButtonProject.BiSAButton cmdAdd 
         Height          =   435
         Left            =   75
         TabIndex        =   42
         Top             =   105
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   767
         Caption         =   "  &Add"
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
         Picture         =   "MstRegisterNasabah.frx":09A3
      End
      Begin BiSAButtonProject.BiSAButton cmdKeluar 
         Cancel          =   -1  'True
         Height          =   435
         Left            =   10590
         TabIndex        =   43
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
         Picture         =   "MstRegisterNasabah.frx":0B4E
      End
   End
End
Attribute VB_Name = "MstRegisterNasabah"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lClick As Boolean
Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.data
Dim lEdit As Boolean
Dim nPos As SisPos
Dim cSQL As String

Private Sub BiSAButton1_Click()
  Load trKonversi
  trKonversi.Show
End Sub

Private Sub cAgama_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "Agama", "Kode,Keterangan", "Kode", sisContent, cAgama.Text)
  cAgama.Text = cAgama.Browse(dbData)
  If Not dbData.eof Then
    cAgama.Text = GetNull(dbData!Kode)
    cNAgama.Text = GetNull(dbData!Keterangan)
  End If
End Sub

Private Sub cCabang_Validate(Cancel As Boolean)
  If nPos = Add Then
    GetNomorRegister False
  End If
End Sub

Private Sub cPekerjaan_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "Pekerjaan", "kode,keterangan", "kode", sisContent, cPekerjaan.Text)
  cPekerjaan.Text = cPekerjaan.Browse(dbData)
  If Not dbData.eof Then
    cPekerjaan.Text = GetNull(dbData!Kode)
    cNPekerjaan.Text = GetNull(dbData!Keterangan)
  End If
End Sub

Private Sub GetRegister()
  cKode.Text = Padl(Trim(cKode.Text), 6, "0")
End Sub

Private Sub cKode_Validate(Cancel As Boolean)
  If cKode.LastKey = 13 Then
      GetRegister
      Set dbData = objData.Browse(GetDSN, "RegisterNasabah", "Kode", "Kode", sisAssign, cCabang.Text & "." & cKode.Text)
      If Not dbData.eof Then
        If nPos = Add Then
          MsgBox "Nomor Register Sudah Ada. Silahkan ulangi pengisian", vbInformation
          Cancel = True
          cKode.Default
          cKode.SetFocus
          Exit Sub
        End If
        GetMemory
        If nPos = Delete Then DeleteData
      ElseIf dbData.eof And nPos <> Add Then
        MsgBox "Data tidak ada. Silahkan ulangi pengisian", vbInformation
        Cancel = True
        initvalue
        cKode.SetFocus
        Exit Sub
      End If
  End If
End Sub

Private Sub GetMemory()
Dim vaJoin
Dim cField As String

  cField = "r.*,a.Keterangan as NamaAgama,p.Keterangan as NamaPekerjaan,"
  cField = cField & " w.Keterangan as NamaWilayah"
  vaJoin = Array("Left Join Agama a on a.kode=r.Agama", _
                 "Left Join Pekerjaan p on p.Kode=r.Pekerjaan", _
                 "Left Join Wilayah w on w.Kode = r.Wilayah")
  Set dbData = objData.Browse(GetDSN, "Registernasabah r", cField, "r.Kode", sisAssign, cCabang.Text & "." & cKode.Text, , , vaJoin)
  If Not dbData.eof Then
    cKode.Text = Mid(GetNull(dbData!Kode), 4)
    SetOpt optJenisAnggota, GetNull(dbData!JenisAnggota)
    dTglRegister.Value = GetNull(dbData!TglRegister)
    cNama.Text = GetNull(dbData!nama)
    cSex.Text = GetNull(dbData!Kelamin)
    optSex(IIf(cSex.Text = "L", 0, 1)).Value = True
    cGolDarah.Text = dbData!GolonganDarah
    Select Case cGolDarah.Text
      Case Is = "A"
        OptGolonganDarah(0).Value = True
      Case Is = "B"
        OptGolonganDarah(1).Value = True
      Case Is = "AB"
        OptGolonganDarah(2).Value = True
      Case Is = "O"
        OptGolonganDarah(3).Value = True
    End Select
    
    cTempatLahir.Text = GetNull(dbData!TempatLahir)
    dTglLahir.Value = GetNull(dbData!TglLahir)
    cStatusPerkawinan.Text = GetNull(dbData!StatusPerkawinan)
    optStatusPerkawinan(IIf(cStatusPerkawinan.Text = "K", 0, 1)).Value = True
    cKTP.Text = GetNull(dbData!KTP)
    cAgama.Text = GetNull(dbData!Agama)
    cNAgama.Text = GetNull(dbData!NamaAgama)
    cPekerjaan.Text = GetNull(dbData!Pekerjaan)
    cNPekerjaan.Text = GetNull(dbData!NamaPekerjaan)
    dTglKTP.Value = GetNull(dbData!TglKtp)
    
    cAlamatRumah.Text = GetNull(dbData!alamat)
    cTeleponRumah.Text = GetNull(dbData!Telepon)
    cWilayah.Text = GetNull(dbData!Wilayah)
    cNamaWilayah.Text = GetNull(dbData!Namawilayah)
    cNPWP.Text = GetNull(dbData!NPWP)
    cAlamatKantor.Text = GetNull(dbData!AlamatKantor)
    cKodePosKantor.Text = GetNull(dbData!KodePosKantor)
    cTeleponKantor.Text = GetNull(dbData!TeleponKantor)
    cFaxKantor.Text = GetNull(dbData!FaxKantor)
    
    Image1.Picture = LoadPicture(GetPicture(GetNull(dbData!Path, "")))
    CommonDialog1.FileName = GetNull(dbData!Path, "")
    
    Image2.Picture = LoadPicture(GetPicture(GetNull(dbData!Path1, "")))
    CommonDialog2.FileName = GetNull(dbData!Path1, "")
  End If
End Sub

Private Sub cmdAdd_Click()
  nPos = Add
  GetEdit True
  initvalue
  cNama.Button = False
  cCabang.SetFocus
End Sub

Private Sub GetEdit(lPar As Boolean)
  lEdit = lPar
  BiSAFrame1.Enabled = lPar
  SetButton cmdSimpan, cmdKeluar, cmdAdd, cmdEdit, cmdHapus, nPos, lPar, cmdAktivasi
  If lPar Then
    If nPos = Add Then
      cKode.Enabled = False
      cKode.BackColor = vbButtonFace
    Else
      cKode.Enabled = True
      cKode.BackColor = vbWindowBackground
      cKode.CaptionBackColor = vbButtonFace
    End If
    'cKode.SetFocus
  End If
End Sub

Private Sub cmdAktivasi_Click()
  frmAktivasi.Action Me
End Sub

Private Sub cmdEdit_Click()
  nPos = Edit
  GetEdit True
  initvalue
  cmdHapus.Enabled = True
  cNama.Button = True
  cCabang.SetFocus
End Sub

Private Sub cmdFoto_Click()
  CommonDialog1.InitDir = aCfg(msPicturePath, App.Path)
  
  CommonDialog1.filter = "Picture (*.BMP;*.JPG;*.GIF) |*.BMP;*.JPG;*.GIF|"
  CommonDialog1.Action = 1
  Image1.Picture = LoadPicture(GetPicture(CommonDialog1.FileName))
  
  If Not CommonDialog1.FileName = "" Then
    UpdCfg msPicturePath, left(CommonDialog1.FileName, RAT("\", CommonDialog1.FileName) - 1)
  End If
End Sub

Private Sub cmdHapus_Click()
  nPos = Delete
  GetEdit True
  initvalue
  cCabang.SetFocus
End Sub

Private Sub DeleteData()
  If MsgBox("Data Benar-benar Dihapus ?", vbYesNo + vbExclamation) = vbYes Then
    objData.Delete GetDSN, "RegisterNasabah", "Kode", sisAssign, cCabang.Text & "." & cKode.Text
  End If
  GetEdit False
  initvalue
End Sub

Private Sub cmdKeluar_Click()
  If Not lEdit Then
    Unload Me
  Else
    initvalue
    GetEdit False
  End If
End Sub

Private Sub cmdSimpan_Click()
Dim vaField, vaValue
Dim cNoRegisterNasabah As String

  If ValidSaving() Then
    If MsgBox("Data benar-benar sudah Valid ?", vbYesNo + vbInformation) = vbYes Then
      
      GetNomorRegister True
      cNoRegisterNasabah = cCabang.Text & "." & cKode.Text
      vaField = Array("Kode", "TglRegister", "Nama", "Kelamin", "GolonganDarah", _
                      "TempatLahir", "TglLahir", "StatusPerkawinan", _
                      "KTP", "Agama", "Pekerjaan", "Wilayah", "Alamat", "Telepon", _
                      "AlamatKantor", "KodePosKantor", "TeleponKantor", "FaxKantor", _
                      "Path", "Path1", "TglKtp", "NPWP", "JenisAnggota")
      vaValue = Array(cNoRegisterNasabah, dTglRegister.Value, cNama.Text, cSex.Text, cGolDarah.Text, _
                      cTempatLahir.Text, dTglLahir.Value, cStatusPerkawinan.Text, _
                      cKTP.Text, cAgama.Text, cPekerjaan.Text, cWilayah.Text, cAlamatRumah.Text, cTeleponRumah.Text, _
                      cAlamatKantor.Text, cKodePosKantor.Text, cTeleponKantor.Text, cFaxKantor.Text, _
                      CommonDialog1.FileName, CommonDialog2.FileName, dTglKTP.Value, cNPWP.Text, GetOpt(optJenisAnggota))
      objData.Update GetDSN, "RegisterNasabah", "Kode='" & cNoRegisterNasabah & "'", vaField, vaValue
      initvalue
      
      GetEdit False
    End If
  End If
End Sub

Static Function ValidSaving() As Boolean
  ValidSaving = True
  
  If Not CheckData(cKode.Text, "Kode Register Nasabah Harus Diisi, Silahkan Mengulangi Pengisian") Then
    ValidSaving = False
    cKode.SetFocus
    Exit Function
  End If
  
  If Not CheckData(cNama.Text, "Nama Register Nasabah Harus Diisi, Silahkan Mengulangi Pengisian") Then
    ValidSaving = False
    cNama.SetFocus
    Exit Function
  End If
  
  If Not CheckData(cTempatLahir.Text, "Tempat Lahir Harus Diisi, Silahkan Mengulangi Pengisian") Then
    ValidSaving = False
    cTempatLahir.SetFocus
    Exit Function
  End If
  
  If Not CheckData(dTglLahir.Value, "Tanggal Lahir Harus Diisi, Silahkan Mengulangi Pengisian") Then
    ValidSaving = False
    dTglLahir.SetFocus
    Exit Function
  End If
  
  If Not CheckData(cKTP.Text, "No KTP Harus Diisi, Silahkan Mengulangi Pengisian") Then
    ValidSaving = False
    cKTP.SetFocus
    Exit Function
  End If

  If Not CheckData(cAgama.Text, "Agama Harus Diisi, Silahkan Mengulangi Pengisian") Then
    ValidSaving = False
    cAgama.SetFocus
    Exit Function
  End If

  If Not CheckData(cPekerjaan.Text, "Pekerjaan Harus Diisi, Silahkan Mengulangi Pengisian") Then
    ValidSaving = False
    cPekerjaan.SetFocus
    Exit Function
  End If

  If Not CheckData(cAlamatRumah.Text, "Alamat Rumah Harus Diisi, Silahkan Mengulangi Pengisian") Then
    ValidSaving = False
    cAlamatRumah.SetFocus
    Exit Function
  End If
  
  If Not CheckData(cWilayah.Text, "Wilayah Harus Diisi, Silahkan Mengulangi Pengisian") Then
    ValidSaving = False
    cWilayah.SetFocus
    Exit Function
  End If
End Function

Private Sub cmdTTD_Click()
  CommonDialog2.InitDir = aCfg(msPicturePath, App.Path)
  
  CommonDialog2.filter = "Picture (*.BMP;*.JPG;*.GIF) |*.BMP;*.JPG;*.GIF|"
  CommonDialog2.Action = 1
  Image2.Picture = LoadPicture(GetPicture(CommonDialog2.FileName))
  
  If Not CommonDialog2.FileName = "" Then
    UpdCfg msPicturePath, left(CommonDialog2.FileName, RAT("\", CommonDialog2.FileName) - 1)
  End If
End Sub

Private Sub cNama_ButtonClick()
  If nPos = Edit Then
    Set dbData = objData.Browse(GetDSN, "RegisterNasabah", "Nama,Alamat,Kode,Path", "Nama", sisContent, cNama.Text, , "Nama")
    cNama.Text = cNama.Browse(dbData, Array("Nama", "Alamat"))
    If Not dbData.eof Then
      cCabang.Text = left(GetNull(dbData!Kode), 2)
      cKode.Text = Mid(GetNull(dbData!Kode), 4)
      GetMemory
    End If
  End If
End Sub

Private Sub initvalue()
  optJenisAnggota(0).Value = True
  cKode.Default
  dTglRegister.Value = Date
  cNama.Default
  optSex(0).Value = True
  cSex.Text = "L"
  OptGolonganDarah(0).Value = True
  cGolDarah.Default
  cTempatLahir.Default
  dTglLahir.Value = Now
  optStatusPerkawinan(0).Value = True
  cStatusPerkawinan.Text = "K"
  cKTP.Default
  dTglKTP.Value = Now
  cAgama.Default
  cNAgama.Default
  cPekerjaan.Default
  cNPekerjaan.Default
  
  cAlamatRumah.Default
  cTeleponRumah.Default
  cWilayah.Default
  cNamaWilayah.Default
  cAlamatKantor.Default
  cKodePosKantor.Default
  cTeleponKantor.Default
  cFaxKantor.Default
  cNPWP.Default
  
  Image1.Picture = LoadPicture("")
  Image2.Picture = LoadPicture("")
  SSTab1.Tab = 0
End Sub

Private Sub cPekerjaan_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    SSTab1.Tab = 1
    cAlamatRumah.SetFocus
  End If
End Sub

Private Sub cWilayah_ButtonClick()
  'Set dbData = objData.Pick(GetDSN, "Wilayah", "Kode", cWilayah, "Kode,Keterangan")
  Set dbData = objData.Browse(GetDSN, "Wilayah", "Kode,Keterangan", "Kode", sisContent, cWilayah.Text)
  cWilayah.Text = cWilayah.Browse(dbData)
  If Not dbData.eof Then
    cNamaWilayah.Text = GetNull(dbData!Keterangan, "")
  End If
End Sub

Private Sub cWilayah_Validate(Cancel As Boolean)
  If cWilayah.LastKey = 13 Then
    cWilayah_ButtonClick
  End If
End Sub

Private Sub Form_Load()
Dim n As Single
  
  CenterForm Me, True
  initvalue
  GetEdit False
  
  cCabang.Text = aCfg(msKodeCabang, "")
  dTglRegister.Value = Date
  
  TabIndex cCabang, n
  TabIndex cKode, n
  TabIndex dTglRegister, n
  TabIndex optJenisAnggota(0), n
  TabIndex optJenisAnggota(1), n
  TabIndex cNama, n
  TabIndex optSex(0), n
  TabIndex optSex(1), n
  TabIndex OptGolonganDarah(0), n
  TabIndex OptGolonganDarah(1), n
  TabIndex OptGolonganDarah(2), n
  TabIndex OptGolonganDarah(3), n
  TabIndex cTempatLahir, n
  TabIndex dTglLahir, n
  TabIndex optStatusPerkawinan(0), n
  TabIndex optStatusPerkawinan(1), n
  TabIndex cKTP, n
  TabIndex dTglKTP, n
  TabIndex cAgama, n
  TabIndex cPekerjaan, n
  
  TabIndex cAlamatRumah, n
  TabIndex cTeleponRumah, n
  TabIndex cWilayah, n
  TabIndex cAlamatKantor, n
  TabIndex cTeleponKantor, n
  TabIndex cFaxKantor, n
  TabIndex cKodePosKantor, n
  TabIndex cNPWP, n
  
  TabIndex cmdAdd, n
  TabIndex cmdEdit, n
  TabIndex cmdHapus, n
  TabIndex cmdSimpan, n
  TabIndex cmdKeluar, n
  TabIndex cmdAktivasi, n
End Sub

Private Sub OptGolonganDarah_Click(Index As Integer)
  Select Case Index
    Case 0
      cGolDarah.Text = "A"
    Case 1
      cGolDarah.Text = "B"
    Case 2
      cGolDarah.Text = "AB"
    Case 3
      cGolDarah.Text = "O"
  End Select
End Sub

Private Sub OptGolonganDarah_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeysA vbKeyTab, True
  End If
End Sub

Private Sub optJenisAnggota_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    SendKeysA vbKeyTab, True
  End If
End Sub

Private Sub optSex_Click(Index As Integer)
  cSex.Text = IIf(Index = 0, "L", "P")
End Sub

Private Sub optSex_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeysA vbKeyTab, True
  End If
End Sub

Private Sub optStatusPerkawinan_Click(Index As Integer)
  cStatusPerkawinan.Text = IIf(Index = 0, "K", "B")
End Sub

Private Sub optStatusPerkawinan_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeysA vbKeyTab, True
  End If
End Sub

Private Sub GetNomorRegister(ByVal lUpdate As Boolean)
  If nPos = Add Then
    cKode.Text = GetLastNomorRegister(cCabang.Text, , lUpdate)
  End If
End Sub
