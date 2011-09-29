Attribute VB_Name = "a_Registry"
Option Explicit

Private Const REG_SZ As Long = &H1
Private Const REG_DWORD As Long = &H4
Private Const REG_OPTION_NON_VOLATILE As Long = 0

Private Const HKEY_CLASSES_ROOT As Long = &H80000000
Private Const HKEY_CURRENT_USER As Long = &H80000001
Private Const HKEY_LOCAL_MACHINE As Long = &H80000002
Private Const HKEY_USERS As Long = &H80000003

Private Const ERROR_SUCCESS As Long = 0
Private Const ERROR_BADDB As Long = 1009
Private Const ERROR_BADKEY As Long = 1010
Private Const ERROR_CANTOPEN As Long = 1011
Private Const ERROR_CANTREAD As Long = 1012
Private Const ERROR_CANTWRITE As Long = 1013
Private Const ERROR_OUTOFMEMORY As Long = 14
Private Const ERROR_INVALID_PARAMETER As Long = 87
Private Const ERROR_ACCESS_DENIED As Long = 5
Private Const ERROR_MORE_DATA As Long = 234
Private Const ERROR_NO_MORE_ITEMS As Long = 259

Private Const KEY_QUERY_VALUE = &H1&
Private Const KEY_CREATE_SUB_KEY = &H4&
Private Const KEY_SET_VALUE = &H2&

Private Declare Function RegCloseKey Lib "advapi32" (ByVal key As Long) As Long

Private Declare Function RegCreateKeyEx Lib "advapi32" Alias "RegCreateKeyExA" (ByVal key As Long, _
                                                                                ByVal subKey As String, _
                                                                                ByVal reserved As Long, _
                                                                                ByVal lpClass As String, _
                                                                                ByVal options As Long, _
                                                                                ByVal samDesired As Long, _
                                                                                ByVal securityAttributes As Long, _
                                                                                ByRef result As Long, _
                                                                                ByRef disposition As Long) As Long

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal key As Long, _
                                                                            ByVal subKey As String, _
                                                                            ByVal options As Long, _
                                                                            ByVal samDesired As Long, _
                                                                            ByRef result As Long) As Long

Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal key As Long, _
                                                                                  ByVal valueName As String, _
                                                                                  ByVal reserved As Long, _
                                                                                  ByRef keyType As Long, _
                                                                                  ByVal valuePtr As Long, _
                                                                                  ByRef valueSize As Long) As Long

Declare Function RegQueryValueExString Lib "advapi32" Alias "RegQueryValueExA" (ByVal key As Long, _
                                                                                ByVal valueName As String, _
                                                                                ByVal reserved As Long, _
                                                                                ByRef keyType As Long, _
                                                                                ByVal value As String, _
                                                                                ByRef valueSize As Long) As Long

Public Declare Function RegSetValueExString Lib "advapi32" Alias "RegSetValueExA" (ByVal key As Long, _
                                                                                   ByVal valueName As String, _
                                                                                   ByVal reserved As Long, _
                                                                                   ByVal keyType As Long, _
                                                                                   ByVal value As String, _
                                                                                   ByVal valueSize As Long) As Long

Declare Function RegDeleteKey Lib "advapi32" Alias "RegDeleteKeyA" (ByVal hKey As Long, _
                                                                    ByVal lpSubKey As String) As Long

'creates new key
Private Sub CreateNewKey(ByVal root As Long, ByVal keyName As String)
    Dim key As Long
    Dim disposition As Long
    
    Dim ret As Long
    ret = RegCreateKeyEx(root, keyName, 0, vbNullString, REG_OPTION_NON_VOLATILE, KEY_CREATE_SUB_KEY, 0, key, disposition)
    If ret <> ERROR_SUCCESS Then
        Echo ">>> CreateNewKey > RegCreateKeyEx failed with return value " & ret
    End If
    
    RegCloseKey key
End Sub


'reads key value
Function QueryValue(ByVal key As Long, ByVal valueName As String, ByRef value As String) As Long
    Dim ret As Long
    Dim keyType As Long
    Dim datasize As Long
    
    'determine the type and size of value to be read
    ret = RegQueryValueEx(key, valueName, 0, keyType, 0, datasize)
    If ret <> ERROR_SUCCESS And ret <> ERROR_MORE_DATA Then
        Echo ">>> QueryValue > RegQueryValueEx failed with return value " & ret
        QueryValue = ret
        Exit Function
    End If
    
    'ensure the value is actually a string
    If keyType <> REG_SZ Then
        Echo ">>> QueryValue > key type is not REG_SZ"
        QueryValue = ERROR_BADKEY
        Exit Function
    End If
    
    'allocate buffer for read-back
    Dim strbuffer As String
    strbuffer = String(datasize, 0)
    
    'read the string
    ret = RegQueryValueExString(key, valueName, 0, keyType, strbuffer, datasize)
    If ret <> ERROR_SUCCESS Then
        Echo ">>> QueryValue > RegQueryValueExString failed with return value " & ret
        QueryValue = ret
        Exit Function
    End If
    
    'set value
    value = Left$(strbuffer, datasize - 1)
    
    QueryValue = ERROR_SUCCESS
End Function


'reads key value
Private Function ReadKey(ByVal root As Long, keyName As String, valueName As String) As String
    Dim key As Long           'handle of opened key
    Dim value As String       'value of queried key
    
    Dim ret As Long
    ret = RegOpenKeyEx(root, keyName, 0, KEY_QUERY_VALUE, key)
    If ret <> ERROR_SUCCESS Then
        'Echo ">>> ReadKey > RegOpenKeyEx failed with return value " & ret
        'this is ok, now we know that the key does not exist
        Exit Function
    End If
    
    ret = QueryValue(key, valueName, value)
    If ret <> ERROR_SUCCESS Then
        Echo ">>> ReadKey > QueryValue failed with return value " & ret
    End If
    
    ret = RegCloseKey(key)
    If ret <> ERROR_SUCCESS Then
        Echo ">>> ReadKey > RegCloseKey failed with return value " & ret
    End If
    
    ReadKey = value
End Function


'sets key value
Private Sub SetKeyValue(root As Long, _
                        sKeyName As String, _
                        sValueName As String, _
                        vValueSetting As String, _
                        lValueType As Long)
    Dim ret As Long
    Dim key As Long
    
    ret = RegOpenKeyEx(root, sKeyName, 0, KEY_SET_VALUE, key)
    If ret <> ERROR_SUCCESS Then
        Echo ">>> SetKeyValue > RegOpenKeyEx failed with return value " & ret
    End If
    
    ret = RegSetValueExString(key, sValueName, 0, REG_SZ, vValueSetting, Len(vValueSetting))
    If ret <> ERROR_SUCCESS Then
        Echo ">>> SetKeyValue > RegSetValueExString failed with return value " & ret
    End If
    
    RegCloseKey key
End Sub


'deletes registry key
Private Sub DeleteKey(ByVal name As String)
    Dim key As Long
    
    Dim ret As Long
    ret = RegOpenKeyEx(HKEY_CLASSES_ROOT, name, 0, KEY_SET_VALUE, key)
    If ret <> ERROR_SUCCESS Then
        Echo ">>> DeleteKey > RegOpenKeyEx failed with return value " & ret
    End If
    
    If key = 0 Then
        Echo ">>> key does not exist"
        Exit Sub
    End If
    
    ret = RegDeleteKey(key, "")
    If ret <> ERROR_SUCCESS Then
        Echo ">>> DeleteKey > RegDeleteKey failed with return value " & ret
    End If
    
    RegCloseKey key
End Sub

'-----------------------------------------------------------------------------------------------------------------------

'adds/removes file association
Public Sub SetFileAssoc(ByRef ext As String, ByVal create As Boolean)
    
    'sanity check
    If Left$(ext, 1) <> "." Then
        MsgBox "Remdul made a stupid error!", vbExclamation
        Exit Sub
    End If
    
    If create Then
        
        Dim bindingName As String
        bindingName = App.Title & ext
        
        Dim exePath As String
        exePath = Chr(34) & App.path & "\" & App.EXEName & ".exe" & Chr(34) & " %1"
        
        Dim description As String
        description = UCase(Right(ext, Len(ext) - 1)) & " File"
        
        'add key for extension
        CreateNewKey HKEY_CLASSES_ROOT, ext
        SetKeyValue HKEY_CLASSES_ROOT, ext, "", bindingName, REG_SZ
        
        'add file handler
        CreateNewKey HKEY_CLASSES_ROOT, bindingName & "\shell\open\command"
        SetKeyValue HKEY_CLASSES_ROOT, bindingName, "", description, REG_SZ
        SetKeyValue HKEY_CLASSES_ROOT, bindingName & "\shell\open\command", "", exePath, REG_SZ
        
    Else
        
        'remove
        If GetFileAssoc(ext) Then
            DeleteKey ext
        End If
        
    End If
End Sub


'returns whether extension is associated with BfMeshView or not
Public Function GetFileAssoc(ByVal ext As String) As Boolean
Dim appname As String
    appname = ReadKey(HKEY_CLASSES_ROOT, ext, "")
    GetFileAssoc = InStr(appname, App.EXEName) > 0
End Function
