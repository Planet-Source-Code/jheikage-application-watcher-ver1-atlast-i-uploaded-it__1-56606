VERSION 5.00
Begin VB.Form FrmLOCK 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   4920
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8130
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   8130
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Image Image1 
      Height          =   1530
      Left            =   30
      Picture         =   "FrmLOCK.frx":0000
      Stretch         =   -1  'True
      Top             =   390
      Width           =   1680
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "APP WATCH 1 by: IEHJSUCKER. MABUHAY NOYPI!"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   1710
      TabIndex        =   1
      Top             =   495
      Width           =   5040
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "THIS PROGRAM IS PROHIBITED TO RUN ONTO YOUR SYSTEM!"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   30
      TabIndex        =   0
      Top             =   15
      Width           =   9945
   End
End
Attribute VB_Name = "FrmLOCK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Sub Form_Load()
SetWindowPos Me.hwnd, -1, 0, 0, Screen.Width, Screen.Height, &H10
End Sub
