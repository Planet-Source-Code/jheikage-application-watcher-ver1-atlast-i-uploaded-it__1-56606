VERSION 5.00
Begin VB.Form FrmAbt 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About App Watch 1"
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5760
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmAbt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   5760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton BtnUnload 
      Caption         =   "OK"
      Height          =   495
      Left            =   4440
      TabIndex        =   1
      Top             =   1725
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Thanx to www.planet-source-code.com, some codes are borrowed over there, i just mixed them up to build this program! "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   120
      TabIndex        =   3
      Top             =   3000
      Width           =   5595
   End
   Begin VB.Image Image2 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Left            =   120
      Picture         =   "FrmAbt.frx":57E2
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   1320
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "App Watch 1 by iehjsucker. Mabuhay NOYPI! e-mail: www.iehjsucker@hotmail.com."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   1875
      TabIndex        =   2
      Top             =   2400
      Width           =   3795
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   5610
      X2              =   1560
      Y1              =   2340
      Y2              =   2340
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      Index           =   0
      X1              =   5610
      X2              =   1575
      Y1              =   2340
      Y2              =   2340
   End
   Begin VB.Label LblX 
      BackStyle       =   0  'Transparent
      Height          =   1620
      Left            =   1755
      TabIndex        =   0
      Top             =   75
      Width           =   3900
   End
   Begin VB.Image Image1 
      Height          =   1530
      Left            =   60
      Picture         =   "FrmAbt.frx":5BC03
      Stretch         =   -1  'True
      Top             =   45
      Width           =   1680
   End
End
Attribute VB_Name = "FrmAbt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub BtnUnload_Click()
Unload Me
End Sub

Private Sub Form_Load()
LblX.Caption = App.FileDescription
End Sub
