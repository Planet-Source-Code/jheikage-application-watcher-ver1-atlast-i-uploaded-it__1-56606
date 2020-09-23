VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmMain 
   Caption         =   "App Watch 1"
   ClientHeight    =   6330
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8145
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6330
   ScaleWidth      =   8145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin AppWatch.cSysTray mysys 
      Left            =   6120
      Top             =   2760
      _ExtentX        =   900
      _ExtentY        =   900
      InTray          =   0   'False
      TrayTip         =   ""
   End
   Begin VB.ListBox applist 
      Height          =   2220
      Left            =   0
      TabIndex        =   3
      Top             =   1200
      Visible         =   0   'False
      Width           =   1620
   End
   Begin VB.Timer MyInt 
      Interval        =   1000
      Left            =   5340
      Top             =   900
   End
   Begin MSComctlLib.StatusBar MyStatus 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   5955
      Width           =   8145
      _ExtentX        =   14367
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   3810
            Picture         =   "FrmMain.frx":57E2
            Text            =   "App Watch 1 "
            TextSave        =   "App Watch 1 "
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   3810
            Text            =   "Total Processes:"
            TextSave        =   "Total Processes:"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Text            =   "0"
            TextSave        =   "0"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList MyImg 
      Left            =   5790
      Top             =   900
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":AFD4
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":D786
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":FB68
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":15E02
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":1B5F4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar MyBar 
      Align           =   1  'Align Top
      Height          =   1050
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   8145
      _ExtentX        =   14367
      _ExtentY        =   1852
      ButtonWidth     =   2143
      ButtonHeight    =   1799
      Appearance      =   1
      Style           =   1
      ImageList       =   "MyImg"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Add/Remove"
            Object.ToolTipText     =   "Add/Remove Prohibited applications"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Run Application"
            Object.ToolTipText     =   "Run Application"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Kill"
            Object.ToolTipText     =   "Kill Process"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Abouts"
            Object.ToolTipText     =   "About App Watch 1"
            ImageIndex      =   5
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.ListBox ProgList 
      BackColor       =   &H00000000&
      Columns         =   3
      ForeColor       =   &H00FFFFFF&
      Height          =   3660
      Left            =   0
      TabIndex        =   0
      Top             =   1080
      Width           =   5325
   End
   Begin VB.Menu Rclick 
      Caption         =   "Rclick"
      Visible         =   0   'False
      Begin VB.Menu rcAW 
         Caption         =   "&Application Watcher"
      End
      Begin VB.Menu rcRun 
         Caption         =   "&Run"
      End
      Begin VB.Menu rcAddRem 
         Caption         =   "&Add/Remove Prohibited Applications"
      End
      Begin VB.Menu brk 
         Caption         =   "-"
      End
      Begin VB.Menu rcAb 
         Caption         =   "About App Watch 1"
         Shortcut        =   {F1}
      End
      Begin VB.Menu rcexit 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Const MAX_PATH& = 260

Private Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szexeFile As String * MAX_PATH
End Type


Private Declare Function TerminateProcess Lib "kernel32" (ByVal ApphProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal blnheritHandle As Long, ByVal dwAppProcessId As Long) As Long
Private Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, lProcessID As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function SHShutDownDialog Lib "shell32" Alias "#60" (ByVal YourGuess As Long) As Long
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long

Public MyPass As String
Public Function MyApp()

    Const PROCESS_ALL_ACCESS = 0
    Dim uProcess As PROCESSENTRY32
    Dim rProcessFound As Long
    Dim hSnapshot As Long
    Dim szExename As String
    Dim exitCode As Long
    Dim myProcess As Long
    Dim AppKill As Boolean
    Dim appCount As Integer
    Dim i As Integer
    On Local Error GoTo Finish
    appCount = 0

    Const TH32CS_SNAPPROCESS As Long = 2&

    uProcess.dwSize = Len(uProcess)
    hSnapshot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
    rProcessFound = ProcessFirst(hSnapshot, uProcess)
    ProgList.Clear

    Do While rProcessFound
        i = InStr(1, uProcess.szexeFile, Chr(0))
        szExename = LCase$(Left$(uProcess.szexeFile, i - 1))
        ProgList.AddItem (szExename)
        If Right$(szExename, Len(myName)) = LCase$(myName) Then

            appCount = appCount + 1
            myProcess = OpenProcess(PROCESS_ALL_ACCESS, False, uProcess.th32ProcessID)
            IdP(ProgList.ListCount - 1) = uProcess.th32ProcessID
            Call CloseHandle(myProcess)
            If ProcessClose.Check_If_Prohibited(szExename, applist, uProcess.th32ProcessID) = True Then
                'remove from list
                ProgList.RemoveItem ProgList.ListCount - 1
            End If
        End If
        rProcessFound = ProcessNext(hSnapshot, uProcess)
    Loop
    Me.MyStatus.Panels(3).Text = ProgList.ListCount
    Call CloseHandle(hSnapshot)
Finish:
End Function


Private Sub Form_Load()
    Disa True
    showme = False
    App.TaskVisible = False
    ListApps applist
    MyApp
    Me.Visible = False
    mysys.InTray = True
    mysys.TrayTip = Me.Caption
    Set mysys.TrayIcon = Me.Icon
    GetRegpass
End Sub

Private Sub Form_LostFocus()
Me.Visible = False
End Sub

Private Sub Form_Resize()
On Error Resume Next
    If Me.Width <= 8265 Then Me.Width = 8265
    If Me.Height <= 6735 Then Me.Height = 6735
    ProgList.Width = ScaleWidth
    ProgList.Height = ScaleHeight - Me.MyBar.Height - MyStatus.Height - 22
    If Me.WindowState = vbMinimized Then Me.Hide
End Sub

Private Sub Form_Unload(Cancel As Integer)
    showme = False
    FrmPassword.Show 1
    If showme = True Then Cancel = 0 Else Cancel = 1
    Disa False
End Sub

Private Sub MyBar_ButtonClick(ByVal Button As MSComctlLib.Button)
 Select Case Button.Index
    Case 1
        'show prohibited apps
        showme = False
        FrmPassword.Show 1
        If showme = True Then
            MyApps.Show 1
            ListApps applist
        End If
    Case 2
        'run application
        Run_App
    Case 3
    'Close application
        If MyStatus.Panels(4).Text = "" Then Exit Sub
        If LCase(MyStatus.Panels(4).Text) = "appwatch.exe" Then Exit Sub
            If MsgBox("Do you really want to terminate this process?", vbYesNo + vbQuestion, "Terminate " & MyStatus.Panels(4).Text) = vbYes Then
            Dim mylong As Long
            mylong = MyStatus.Panels(5).Text
            Process_Kill IdP(mylong)
        End If
    Case 5
        FrmAbt.Show 1
End Select
End Sub

Private Sub MyInt_Timer()
    MyStatus.Panels(3).Text = "Retreiving processess..."
    MyApp
End Sub

Private Sub mysys_MouseDown(Button As Integer, Id As Long)
    If Button = 2 Then
        'show menus
        PopupMenu Rclick
    End If
End Sub

Private Sub Proglist_Click()
    MyStatus.Panels(4).Text = ProgList.Text
    MyStatus.Panels(5).Text = ProgList.ListIndex
End Sub

Function Run_App()
    On Error GoTo MayError:
        Dim apprun As String
        apprun = InputBox("Enter new application to run:")
        If Len(Trim(apprun)) = 0 Then Exit Function
        If IsNumeric(apprun) Then Exit Function
        Shell apprun, vbNormalFocus
        Exit Function
MayError:
        MsgBox Err.Description
        Exit Function
End Function


Function GetRegpass()
    MyPass = GetSetting("axp", "pxa", "msvba_critical_patch", "monitor")
    prevscr = GetSetting("axp", "pxa", "msvba_critical_patch0", 1)
    Autorun = GetSetting("axp", "pxa", "msvba_critical_patch1", 1)
    Dim xx As New MyReg
    If Autorun = 1 Then
        xx.SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\MICROSOFT\WINDOWS\CURRENTVERSION\RUN", "App Watch", App.Path & "\" & App.EXEName & ".exe"
    Else
        xx.DeleteKeyValue "HKEY_LOCAL_MACHINE\SOFTWARE\MICROSOFT\WINDOWS\CURRENTVERSION\RUN", "App Watch"
    End If
    Set xx = Nothing
End Function

Private Sub rcAb_Click()
    FrmAbt.Show 1
End Sub

Private Sub rcAddRem_Click()
    MyBar_ButtonClick MyBar.Buttons(1)
End Sub

Private Sub rcAW_Click()
    Me.Show
    Me.Visible = True
    Me.WindowState = vbNormal
    Me.Show
End Sub

Private Sub rcexit_Click()
    showme = False
    Me.Show
    Me.Visible = True
    Me.WindowState = vbNormal
    FrmPassword.Show 1
    If showme = True Then Unload Me
End Sub

Private Sub rcRun_Click()
MyBar_ButtonClick MyBar.Buttons(2)
End Sub

Function Disa(bx As Boolean)
SystemParametersInfo 97, bx, CStr(1), &H2
End Function
