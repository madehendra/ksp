VERSION 5.00
Object = "{55473EAC-7715-4257-B5EF-6E14EBD6A5DD}#1.0#0"; "vbalProgBar6.ocx"
Begin VB.Form frmProgress 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1005
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4245
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1005
   ScaleWidth      =   4245
   StartUpPosition =   3  'Windows Default
   Begin vbalProgBarLib6.vbalProgressBar vbalProgressBar1 
      Height          =   315
      Left            =   135
      TabIndex        =   0
      Top             =   525
      Width           =   3915
      _ExtentX        =   6906
      _ExtentY        =   556
      Picture         =   "frmProgress.frx":0000
      ForeColor       =   0
      BarPicture      =   "frmProgress.frx":001C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      XpStyle         =   -1  'True
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   375
      Left            =   150
      TabIndex        =   1
      Top             =   150
      Width           =   3915
   End
End
Attribute VB_Name = "frmProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

