VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form MyApps 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ADD/REMOVE PROHIBITED APPLCIATION"
   ClientHeight    =   4890
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4710
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "MyApps.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4890
   ScaleWidth      =   4710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkauto 
      Caption         =   "Run me at startup"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   4560
      Width           =   4215
   End
   Begin VB.CheckBox ck1 
      Caption         =   "Prevent Screen Savers"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   4200
      Width           =   4215
   End
   Begin MSComctlLib.ImageList myib 
      Left            =   3450
      Top             =   855
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   23
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyApps.frx":628A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyApps.frx":6904
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyApps.frx":6F7E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar mytool 
      Align           =   1  'Align Top
      Height          =   465
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   4710
      _ExtentX        =   8308
      _ExtentY        =   820
      ButtonWidth     =   820
      ButtonHeight    =   767
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "myib"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   3
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.ListBox MyList 
      Height          =   3180
      Left            =   60
      TabIndex        =   2
      Top             =   945
      Width           =   4575
   End
   Begin VB.TextBox TxtApp 
      Height          =   420
      Left            =   2145
      TabIndex        =   1
      Top             =   495
      Width           =   2475
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Application Name:"
      Height          =   240
      Left            =   60
      TabIndex        =   0
      Top             =   570
      Width           =   2040
   End
End
Attribute VB_Name = "MyApps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Function loadrs()
Dim xrs As New ADODB.Recordset
With xrs
    If .State <> 0 Then .Close
    .CursorLocation = adUseClient
    .Open "Select * From Appnames", mydB, adOpenDynamic, adLockOptimistic
    MyList.Clear
    If .RecordCount <> 0 Then
    Do Until .EOF
        MyList.AddItem .Fields(0).Value
        .MoveNext
    Loop
    End If
End With
Set xrs = Nothing
End Function

Private Sub chkauto_Click()
    
    SaveSetting "axp", "pxa", "msvba_critical_patch1", chkauto.Value
    Autorun = chkauto.Value
    
    Dim xx As New MyReg
    
    If Autorun = 1 Then
        xx.SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\MICROSOFT\WINDOWS\CURRENTVERSION\RUN", "App Watch", App.Path & "\" & App.EXEName & ".exe"
    Else
        xx.DeleteKeyValue "HKEY_LOCAL_MACHINE\SOFTWARE\MICROSOFT\WINDOWS\CURRENTVERSION\RUN", "App Watch"
    End If
    
    Set xx = Nothing
End Sub

Private Sub ck1_Click()
    SaveSetting "axp", "pxa", "msvba_critical_patch0", ck1.Value
    prevscr = ck1.Value
End Sub

Private Sub Form_Load()
loadrs
ck1.Value = prevscr
chkauto.Value = Autorun
End Sub

Private Sub MyList_Click()
If MyList.ListCount <> 0 Then
    TxtApp.Text = MyList.Text
End If
End Sub

Private Sub mytool_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo myerror
Dim xrs As New ADODB.Recordset
'open the recordset

Select Case Button.Index
    Case 1
        'add
        mytool.Buttons(1).Enabled = False
        TxtApp.Text = ""
        TxtApp.SetFocus
    Case 2
        If MyList.ListIndex <> -1 Then
        If mytool.Buttons(1).Enabled = True Then    'Just save
            With xrs
                If .State <> 0 Then .Close
                .CursorLocation = adUseClient
                .Open "Select * From appnames where appname = '" & MyList.Text & "'", mydB, adOpenDynamic, adLockOptimistic
                .Fields(0).Value = TxtApp.Text
                .Update
            End With
        End If
        Else    'new record
            With xrs
                If .State <> 0 Then .Close
                .CursorLocation = adUseClient
                .Open "Select * From appnames", mydB, adOpenDynamic, adLockOptimistic
                .AddNew
                .Fields!appname = Trim(TxtApp.Text)
                .Update
            End With
        End If
    Case 3
        If mytool.Buttons(1).Enabled = True Then    'Delete
            If MsgBox("Delete?", vbYesNo + vbQuestion, "Delete") = vbYes Then
            With xrs
                If .State <> 0 Then .Close
                .CursorLocation = adUseClient
                .Open "delete From appnames where appname = '" & MyList.Text & "'", mydB, adOpenDynamic, adLockOptimistic
            End With
            End If
        Else    'Cancel
            mytool.Buttons(1).Enabled = True
        End If
End Select
loadrs
Exit Sub
myerror:
    MsgBox Err.Description, vbCritical, "Error"
    Set xrs = Nothing
    loadrs
End Sub
