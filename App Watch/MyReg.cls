VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MyReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Type FILETIME
    lLowDateTime    As Long
    lHighDateTime   As Long
End Type
 
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As Any, phkResult As Long, lplDisposition As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Private Declare Function RegQueryValueExA Lib "advapi32.dll" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByRef lpData As Long, lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Private Declare Function RegSetValueExA Lib "advapi32.dll" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByRef lpData As Long, ByVal cbData As Long) As Long
Private Declare Function RegSetValueExB Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByRef lpData As Byte, ByVal cbData As Long) As Long


Private Const ERROR_SUCCESS = 0&
Private Const ERROR_BADDB = 1009&
Private Const ERROR_BADKEY = 1010&
Private Const ERROR_CANTOPEN = 1011&
Private Const ERROR_CANTREAD = 1012&
Private Const ERROR_CANTWRITE = 1013&
Private Const ERROR_OUTOFMEMORY = 14&
Private Const ERROR_INVALID_PARAMETER = 87&
Private Const ERROR_ACCESS_DENIED = 5&
Private Const ERROR_NO_MORE_ITEMS = 259&
Private Const ERROR_MORE_DATA = 234&

Private Const REG_NONE = 0&
Private Const REG_SZ = 1&
Private Const REG_EXPAND_SZ = 2&
Private Const REG_BINARY = 3&
Private Const REG_DWORD = 4&
Private Const REG_DWORD_LITTLE_ENDIAN = 4&
Private Const REG_DWORD_BIG_ENDIAN = 5&
Private Const REG_LINK = 6&
Private Const REG_MULTI_SZ = 7&
Private Const REG_RESOURCE_LIST = 8&
Private Const REG_FULL_RESOURCE_DESCRIPTOR = 9&
Private Const REG_RESOURCE_REQUIREMENTS_LIST = 10&

Private Const KEY_QUERY_VALUE = &H1&
Private Const KEY_SET_VALUE = &H2&
Private Const KEY_CREATE_SUB_KEY = &H4&
Private Const KEY_ENUMERATE_SUB_KEYS = &H8&
Private Const KEY_NOTIFY = &H10&
Private Const KEY_CREATE_LINK = &H20&
Private Const READ_CONTROL = &H20000
Private Const WRITE_DAC = &H40000
Private Const WRITE_OWNER = &H80000
Private Const SYNCHRONIZE = &H100000
Private Const STANDARD_RIGHTS_REQUIRED = &HF0000
Private Const STANDARD_RIGHTS_READ = READ_CONTROL
Private Const STANDARD_RIGHTS_WRITE = READ_CONTROL
Private Const STANDARD_RIGHTS_EXECUTE = READ_CONTROL
Private Const KEY_READ = STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY
Private Const KEY_WRITE = STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY
Private Const KEY_EXECUTE = KEY_READ

Private hKey As Long, MainKeyHandle As Long
Private rtn As Long, lBuffer As Long, sBuffer As String
Private lBufferSize As Long
Private lDataSize As Long
Private ByteArray() As Byte

Private DisplayErrorMsg As Boolean

Private Sub Class_Initialize()

    DisplayErrorMsg = False
    
End Sub


Public Property Let SetDisplayErrorMsg(vNewValue As Variant)

    DisplayErrorMsg = vNewValue

End Property


Public Function SetDWordValue(ByVal sKey As String, ByVal sKeyName As String, ByVal KeyValue As Long)

SetDWordValue = False
Call ParseKey(sKey, MainKeyHandle)

If MainKeyHandle Then
   rtn = RegOpenKeyEx(MainKeyHandle, sKey, 0, KEY_WRITE, hKey)
   If rtn = ERROR_SUCCESS Then
      rtn = RegSetValueExA(hKey, sKeyName, 0, REG_DWORD, KeyValue, 4)
      If Not rtn = ERROR_SUCCESS Then
         If DisplayErrorMsg = True Then
            MsgBox GetErrorMsg(rtn)
         End If
      Else
         SetDWordValue = True
      End If
      rtn = RegCloseKey(hKey)
   Else
      If DisplayErrorMsg = True Then
         MsgBox GetErrorMsg(rtn)
      End If
   End If
End If

End Function


Public Function GetDWordValue(ByVal sKey As String, ByVal sKeyName As String)

Call ParseKey(sKey, MainKeyHandle)

If MainKeyHandle Then
   rtn = RegOpenKeyEx(MainKeyHandle, sKey, 0, KEY_READ, hKey)
   If rtn = ERROR_SUCCESS Then
      rtn = RegQueryValueExA(hKey, sKeyName, 0, REG_DWORD, lBuffer, 4)
      If rtn = ERROR_SUCCESS Then
         rtn = RegCloseKey(hKey)
         GetDWordValue = lBuffer
      Else
         GetDWordValue = "Error"
         If DisplayErrorMsg = True Then
            MsgBox GetErrorMsg(rtn)
         End If
      End If
   Else
      GetDWordValue = "Error"
      If DisplayErrorMsg = True Then
         MsgBox GetErrorMsg(rtn)
      End If
   End If
End If

End Function


Public Function SetBinaryValue(ByVal sKey As String, ByVal sKeyName As String, KeyValue As String)

SetBinaryValue = False
Call ParseKey(sKey, MainKeyHandle)

If MainKeyHandle Then
   rtn = RegOpenKeyEx(MainKeyHandle, sKey, 0, KEY_WRITE, hKey)
   If rtn = ERROR_SUCCESS Then
      lDataSize = Len(KeyValue)
      ReDim ByteArray(lDataSize)
      For i = 1 To lDataSize
      ByteArray(i) = Asc(Mid$(KeyValue, i, 1))
      Next
      rtn = RegSetValueExB(hKey, sKeyName, 0, REG_BINARY, ByteArray(1), lDataSize)
      If Not rtn = ERROR_SUCCESS Then
         If DisplayErrorMsg = True Then
            MsgBox GetErrorMsg(rtn)
         End If
      Else
         SetBinaryValue = True
      End If
      rtn = RegCloseKey(hKey)
   Else
      If DisplayErrorMsg = True Then
         MsgBox GetErrorMsg(rtn)
      End If
   End If
End If

End Function


Public Function GetBinaryValue(ByVal sKey As String, ByVal sKeyName As String)

Call ParseKey(sKey, MainKeyHandle)

If MainKeyHandle Then
   rtn = RegOpenKeyEx(MainKeyHandle, sKey, 0, KEY_READ, hKey)
   If rtn = ERROR_SUCCESS Then
      lBufferSize = 1
      rtn = RegQueryValueEx(hKey, sKeyName, 0, REG_BINARY, 0, lBufferSize)
      sBuffer = Space(lBufferSize)
      rtn = RegQueryValueEx(hKey, sKeyName, 0, REG_BINARY, sBuffer, lBufferSize)
      If rtn = ERROR_SUCCESS Then
         rtn = RegCloseKey(hKey)
         GetBinaryValue = sBuffer
      Else
         GetBinaryValue = "Error"
         If DisplayErrorMsg = True Then
            MsgBox GetErrorMsg(rtn)
         End If
      End If
   Else
      GetBinaryValue = "Error"
      If DisplayErrorMsg = True Then
         MsgBox GetErrorMsg(rtn)
      End If
   End If
End If

End Function


Public Function SetStringValue(ByVal sKey As String, ByVal sKeyName As String, ByVal KeyValue As String)

SetStringValue = False
Call ParseKey(sKey, MainKeyHandle)

If MainKeyHandle Then
   rtn = RegOpenKeyEx(MainKeyHandle, sKey, 0, KEY_WRITE, hKey)
   If rtn = ERROR_SUCCESS Then
      rtn = RegSetValueEx(hKey, sKeyName, 0, REG_SZ, ByVal KeyValue, Len(KeyValue))
      If Not rtn = ERROR_SUCCESS Then
         If DisplayErrorMsg = True Then
            MsgBox GetErrorMsg(rtn)
         End If
      Else
         SetStringValue = True
      End If
      rtn = RegCloseKey(hKey)
   Else
      If DisplayErrorMsg = True Then
         MsgBox GetErrorMsg(rtn)
      End If
   End If
End If

End Function


Public Function GetStringValue(ByVal sKey As String, ByVal sKeyName As String)

lBufferSize = 0
sBuffer = ""

Call ParseKey(sKey, MainKeyHandle)

If MainKeyHandle Then
   rtn = RegOpenKeyEx(MainKeyHandle, sKey, 0, KEY_READ, hKey)
   If rtn = ERROR_SUCCESS Then
      sBuffer = Space(255)
      lBufferSize = Len(sBuffer)
      rtn = RegQueryValueEx(hKey, sKeyName, 0, REG_SZ, sBuffer, lBufferSize)
      If rtn = ERROR_SUCCESS Then
         rtn = RegCloseKey(hKey)
         sBuffer = Trim(sBuffer)
         GetStringValue = Left(sBuffer, lBufferSize - 1)
      Else
         GetStringValue = "Error"
         If DisplayErrorMsg = True Then
            MsgBox GetErrorMsg(rtn)
         End If
      End If
   Else
      GetStringValue = "Error"
      If DisplayErrorMsg = True Then
         MsgBox GetErrorMsg(rtn)
      End If
   End If
End If
MsgBox GetStringValue
End Function


Public Function CreateKey(ByVal sKey As String)

    CreateKey = False
    Call ParseKey(sKey, MainKeyHandle)
    
    If MainKeyHandle Then
       rtn = RegCreateKey(MainKeyHandle, sKey, hKey)
       If rtn = ERROR_SUCCESS Then
          rtn = RegCloseKey(hKey)
          CreateKey = True
       End If
    End If

End Function


Public Function DeleteKey(ByVal Keyname As String)

    DeleteKey = False
    Call ParseKey(Keyname, MainKeyHandle)
    
    If MainKeyHandle Then
        rtn = RegDeleteKey(MainKeyHandle, Keyname)
        If (rtn <> ERROR_SUCCESS) Then
            If DisplayErrorMsg = True Then
                MsgBox GetErrorMsg(rtn)
            End If
        Else
            DeleteKey = True
        End If
    End If
    
End Function


Public Function DeleteKeyValue(ByVal sKeyName As String, ByVal sValueName As String)

    DeleteKeyValue = False
    Dim hKey As Long

    Call ParseKey(sKeyName, MainKeyHandle)

    If MainKeyHandle Then

        rtn = RegOpenKeyEx(MainKeyHandle, sKeyName, 0, KEY_WRITE, hKey)
        If (rtn = ERROR_SUCCESS) Then
            rtn = RegDeleteValue(hKey, sValueName)
            If (rtn <> ERROR_SUCCESS) Then
                If DisplayErrorMsg = True Then
                    MsgBox GetErrorMsg(rtn)
                End If
            Else
                DeleteKeyValue = True
            End If
            rtn = RegCloseKey(hKey)

        End If

    End If

End Function


Public Function DeleteAllKeySubItems()

    DeleteAllKeySubItems = False

End Function


Public Function KeyExist(ByVal sKey As String)
    Dim hKey As Long

    Call ParseKey(sKey, MainKeyHandle)

    If MainKeyHandle Then
        rtn = RegOpenKeyEx(MainKeyHandle, sKey, 0, KEY_READ, hKey)
        If rtn = ERROR_SUCCESS Then
            KeyExist = True
        Else
            KeyExits = False
        End If
    End If
    
End Function


Public Function KeyValueExist(ByVal sKey As String, ByVal sKeyName As String)
    Dim hKey As Long
    Dim lActualType As Long
    Dim lSize As Long
    
    Dim sTmp As String

    Call ParseKey(sKey, MainKeyHandle)

    If MainKeyHandle Then
        
        rtn = RegOpenKeyEx(MainKeyHandle, sKey, 0, KEY_READ, hKey)
        If (rtn = ERROR_SUCCESS) Then
            
            rtn = RegQueryValueEx(hKey, ByVal sKeyName, 0&, lActualType, sTmp, lSize)
            If (rtn = ERROR_SUCCESS) Then
                KeyValueExist = True
            Else
                KeyValueExist = False
            End If
        
        End If
        
    End If

End Function


Private Sub ParseKey(Keyname As String, Keyhandle As Long)
    
rtn = InStr(Keyname, "\")

If Left(Keyname, 5) <> "HKEY_" Or Right(Keyname, 1) = "\" Then
   MsgBox "Incorrect Format:" + Chr(10) + Chr(10) + Keyname
   Exit Sub
ElseIf rtn = 0 Then
   Keyhandle = GetMainKeyHandle(Keyname)
   Keyname = ""
Else
   Keyhandle = GetMainKeyHandle(Left(Keyname, rtn - 1))
   Keyname = Right(Keyname, Len(Keyname) - rtn)
End If

End Sub


Private Function GetMainKeyHandle(MainKeyName As String) As Long

Const HKEY_CLASSES_ROOT = &H80000000
Const HKEY_CURRENT_USER = &H80000001
Const HKEY_LOCAL_MACHINE = &H80000002
Const HKEY_USERS = &H80000003
Const HKEY_PERFORMANCE_DATA = &H80000004
Const HKEY_CURRENT_CONFIG = &H80000005
Const HKEY_DYN_DATA = &H80000006
   
Select Case MainKeyName
       Case "HKEY_CLASSES_ROOT"
            GetMainKeyHandle = HKEY_CLASSES_ROOT
       Case "HKEY_CURRENT_USER"
            GetMainKeyHandle = HKEY_CURRENT_USER
       Case "HKEY_LOCAL_MACHINE"
            GetMainKeyHandle = HKEY_LOCAL_MACHINE
       Case "HKEY_USERS"
            GetMainKeyHandle = HKEY_USERS
       Case "HKEY_PERFORMANCE_DATA"
            GetMainKeyHandle = HKEY_PERFORMANCE_DATA
       Case "HKEY_CURRENT_CONFIG"
            GetMainKeyHandle = HKEY_CURRENT_CONFIG
       Case "HKEY_DYN_DATA"
            GetMainKeyHandle = HKEY_DYN_DATA
End Select

End Function

Private Function GetErrorMsg(lErrorCode As Long) As String

Select Case lErrorCode
       Case 1009, 1015
            GetErrorMsg = "The Registry Database is corrupt!"
       Case 2, 1010
            GetErrorMsg = "Bad Key Name"
       Case 1011
            GetErrorMsg = "Can't Open Key"
       Case 4, 1012
            GetErrorMsg = "Can't Read Key"
       Case 5
            GetErrorMsg = "Access to this key is denied"
       Case 1013
            GetErrorMsg = "Can't Write Key"
       Case 8, 14
            GetErrorMsg = "Out of memory"
       Case 87
            GetErrorMsg = "Invalid Parameter"
       Case 234
            GetErrorMsg = "There is more data than the buffer has been allocated to hold."
       Case Else
            GetErrorMsg = "Undefined Error Code:  " & Str$(lErrorCode)
End Select

End Function



