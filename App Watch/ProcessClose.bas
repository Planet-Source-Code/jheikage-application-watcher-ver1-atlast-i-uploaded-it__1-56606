Attribute VB_Name = "ProcessClose"
Option Explicit

Public showme As Boolean
Public prevscr As Integer, Autorun As Integer
Public mydB As New ADODB.Connection
Public MyRs As New ADODB.Recordset

Public Const PROCESS_QUERY_INFORMATION As Long = &H400
Public Const PROCESS_TERMINATE As Long = &H1

Private Declare Function OpenProcess Lib "kernel32.dll" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Boolean, ByVal dwProcessId As Long) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32.dll" (ByVal hProcess As Long, ByRef lpExitCode As Long) As Boolean
Private Declare Function TerminateProcess Lib "kernel32.dll" (ByVal hProcess As Long, ByVal uExitCode As Long) As Boolean
Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Boolean

Public IdP(0 To 9999) As Long

Const sLocation As String = "App watch process"

Public Sub File_WriteTo(Text As String)
    Open App.Path & "\LOG\APPWATCH.RDX" For Append As #1
        Print #1, Text
    Close #1
End Sub

Public Sub MayError(numero As Long, Katangian As String, froms As String, who As String)
    File_WriteTo "Mga Error: " & numero & " sa " & froms & " Galing sa " & who & " >>> " & Katangian
End Sub

Public Sub Process_Kill(P_ID As Long)
    '// Kill the wanted process
    
    Dim hProcess As Long
    Dim lExitCode As Long
    
    hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_TERMINATE, False, P_ID): If hProcess = 0 Then Call MayError(Err.LastDllError, "OpenProcess failed", sLocation, "Kill_Process")
    
    If GetExitCodeProcess(hProcess, lExitCode) = False Then Call MayError(Err.LastDllError, "ExitCode failed", sLocation, "Process Killing")
    If TerminateProcess(hProcess, lExitCode) = False Then Call MayError(Err.LastDllError, "Terminate failed", sLocation, "Process Killing")
    
    If CloseHandle(hProcess) = False Then Call MayError(Err.LastDllError, "Handle failed", sLocation, "Process Killing")
End Sub

'The database of prohibited process
Public Function connectAPPLISTS() As Boolean
On Error GoTo MyErb
With mydB
    If .State <> 0 Then .Close
    .CursorLocation = adUseClient
    .Open "Provider=Microsoft.jet.oledb.4.0;Persist security info=false;JET OLEDB:database Password=vip;Data source=" & App.Path & "\probapp\applists.dbx"
End With
connectAPPLISTS = True
Exit Function
MyErb:
    MsgBox "Cannot connect to your prohibited applications. This feature is currently not available to the system.", vbInformation, "Prohibited applications Not Active"
    connectAPPLISTS = False
    FrmMain.MyBar.Buttons(1).Enabled = False
End Function

Public Function ListApps(Lst As ListBox)
    On Error GoTo Erb
    Lst.Clear
    If connectAPPLISTS = True Then
        With MyRs
            If .State <> 0 Then .Close
            .CursorLocation = adUseClient
            .Open "Select * From Appnames", mydB, adOpenDynamic, adLockOptimistic
            If .RecordCount <> 0 Then
                Do Until .EOF
                    Lst.AddItem .Fields(0).Value
                    .MoveNext
                Loop
            End If
        End With
    End If
    Exit Function
Erb:
MsgBox "Cannot connect to your prohibited applications. This feature is currently not available to the system.", vbInformation, "Prohibited applications Not Active"
Lst.Clear
End Function

Public Function Check_If_Prohibited(appname As String, Lst As ListBox, Id As Long) As Boolean
    If Lst.ListCount <> 0 Then
        Dim i As Integer
        For i = 0 To Lst.ListCount - 1
            If UCase(appname) = UCase(Lst.List(i)) Then
                'Terminate
                Process_Kill Id
                FrmLOCK.Show
                MsgBox "This application is currently prohibited to run. App Watch is terminating this process.", vbInformation + vbSystemModal, "Application is Prohibited"
                Unload FrmLOCK
                File_WriteTo appname & " ay na-terminate na kasi ito ay PROHIBITED APPLICATION."
                Check_If_Prohibited = True
                Exit Function
            End If
            'check screen savers
            Dim j As Integer
            If prevscr = 1 Then
                If InStr(1, appname, ".scr", vbTextCompare) <> 0 Then
                    'disable screen saver
                    Process_Kill Id
                    FrmLOCK.Show
                    MsgBox "This application is currently prohibited to run. App Watch is terminating this process.", vbInformation + vbSystemModal, "Application is Prohibited"
                    Unload FrmLOCK
                    File_WriteTo appname & " ay na-terminate na kasi ito ay PROHIBITED APPLICATION."
                    Check_If_Prohibited = True
                    Exit Function
                End If
            End If
        Next
    End If
    Check_If_Prohibited = False
End Function
'end

