Attribute VB_Name = "a_Misc"
Option Explicit

Public uveditor_isloaded As Boolean

Public Type color3f
    r As Single
    g As Single
    b As Single
End Type
Public Type color4f
    r As Single
    g As Single
    b As Single
    a As Single
End Type
Public Const maxcolors As Long = 20
Public colortable(0 To maxcolors) As color4f


Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal dst As Long, ByVal src As Long, ByVal size As Long)

Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Long) As Long
Private Declare Function lstrcpyn Lib "kernel32" Alias "lstrcpynA" (ByVal Recipient As String, _
                                                                    ByVal pSourceString As Long, _
                                                                    ByVal iMaxLength As Long) As Long

'converts character array to VB string in an unsafe way
Public Function CharToString(ByVal address As Long, Optional ByVal num As Long) As String
Dim n As Long
    If num > 0 Then
        n = num
    Else
        n = lstrlen(address)
    End If
    CharToString = String(n, Chr(32))
    lstrcpyn CharToString, address, n + 1
End Function


'converts character array to VB string in a safe way
Public Function SafeString(ByRef buffer() As Byte, ByVal num As Long) As String
    Dim i As Long
    For i = 0 To num - 1
        If buffer(i) = 0 Then
            Exit Function
        Else
            If buffer(i) > 31 And buffer(i) < 127 Then
                SafeString = SafeString & Chr$(buffer(i))
            Else
                SafeString = SafeString & "?"
            End If
        End If
    Next i
End Function


'converts 16 bit signed short to 32 bit unsigned int
Function ShortToInt(ByRef v As Integer) As Long
    If v < 0 Then
        ShortToInt = v + 65536
    Else
        ShortToInt = v
    End If
End Function


'returns boolean as string
Public Function YesNo(ByVal v As Boolean) As String
    If v Then
        YesNo = "Yes"
    Else
        YesNo = "No"
    End If
End Function


'returns directory string from file path string
Public Function GetFilePath(ByVal str As String) As String
Dim pos As Long
Dim pos2 As Long
    pos = InStrRev(str, "\")
    pos2 = InStrRev(str, "/")
    If pos2 > pos Then pos = pos2
    GetFilePath = Left(str, pos)
End Function


'returns file name from file path string
Public Function GetFileName(ByVal str As String) As String
Dim pos As Long
    pos = InStrRev(str, "\")
    If pos = 0 Then pos = InStrRev(str, "/")
    GetFileName = Mid(str, pos + 1, Len(str))
End Function


'returns filename from file path string (note: does same as above, should merge these sometime...)
Public Function GetFilenameFromPath(ByVal str As String) As String
Dim s As Long
    s = InStrRev(str, "\")
    If s = 0 Then s = InStrRev(str, "/")
    GetFilenameFromPath = Right$(str, Len(str) - s)
End Function


'returns name from file path string
Public Function GetNameFromFileName(ByVal str As String) As String
Dim s As Long
Dim e As Long
    s = InStrRev(str, "\")
    If s = 0 Then s = InStrRev(str, "/")
    e = InStrRev(str, ".") - 1
    If e < 0 Then e = 0
    GetNameFromFileName = Mid(str, s + 1, e - s)
End Function


'returns extension string from file path string
Public Function GetFileExt(ByVal str As String) As String
Dim dot As Long
    dot = InStrRev(str, ".")
    'GetFileExt = Right$(str, Len(str) - dot)
    GetFileExt = Mid(str, dot + 1, Len(str))
End Function


'returns whether file exists or not
Public Function FileExist(ByRef filename As String) As Boolean
    On Error GoTo errorhandler
    
    If Len(filename) = 0 Then Exit Function
    If dir$(filename) = "" Then Exit Function
    
    FileExist = True
    
    Exit Function
errorhandler:
End Function


'converts fixed size string into dynamic string
Public Function safeStr(ByRef src As String) As String
Dim p As Long
    p = InStr(1, src, vbNullChar)  'todo: buggy, may crash!!
    If p Then
        safeStr = Left(src, p - 1)
    End If
End Function


'prints text to log
Public Sub Echo(ByVal str As String)
    frmMain.txtLog.Text = frmMain.txtLog.Text & vbCrLf & str
End Sub


'sets status bar text
Public Sub SetStatus(ByVal key As String, ByRef str As String)
    On Error GoTo errhandler
    frmMain.stsMain.Panels(key).Text = str
    'DoEvents
    Exit Sub
errhandler:
    MsgBox "SetStatus" & vbLf & err.description, vbCritical
End Sub


'returns whether app is running in IDE
Public Function IsIdeMode() As Boolean
    On Error GoTo errorhandler
    IsIdeMode = False
    Debug.Print 1 / 0 'division by zero to trigger error
    Exit Function
errorhandler:
    IsIdeMode = True
End Function


'Public Sub SetLeftBits(ByRef dst As Long, ByVal val As Integer)
'    CopyMemory VarPtr(dst), VarPtr(val), 2
'End Sub
'Public Sub SetRightBits(ByRef dst As Long, ByVal val As Integer)
'    CopyMemory VarPtr(dst) + 2, VarPtr(val), 2
'End Sub
'Public Function GetLeftBits(ByRef src As Long) As Integer
'    CopyMemory VarPtr(GetLeftBits), VarPtr(src), 2
'End Function
'Public Function GetRightBits(ByRef src As Long) As Integer
'    CopyMemory VarPtr(GetRightBits), VarPtr(src) + 2, 2
'End Function

Public Sub SetBit(ByRef dst As Long, ByVal pos As Long, ByVal val As Byte)
    CopyMemory VarPtr(dst) + pos, VarPtr(val), 1
End Sub
Public Function GetBit(ByRef src As Long, ByVal pos As Long) As Byte
    CopyMemory VarPtr(GetBit), VarPtr(src) + pos, 1
End Function

Public Function MakeTag(ByVal a As Byte, ByVal b As Byte, ByVal c As Byte) As Long
    'SetLeftBits MakeTag, a
    'SetRightBits MakeTag, b
    MakeTag = 0
    SetBit MakeTag, 0, a
    SetBit MakeTag, 1, b
    SetBit MakeTag, 2, c
End Function


'properly formats floating point number
Public Function fff(ByVal v As Single) As String
    fff = Round(v, 6)
End Function


'generates color table
Public Sub GenColorTable()
Dim i As Long
    
    'generate color table
    Randomize
    For i = 0 To maxcolors
        
        colortable(i).r = 0.25 + (Rnd * 0.75)
        colortable(i).g = 0.25 + (Rnd * 0.75)
        colortable(i).b = 0.25 + (Rnd * 0.75)
        colortable(i).a = 0.75
    Next i
    
    'good set of colors
    Dim b As Single
    b = 0.1
    
    colortable(4).r = b + 225 / 255
    colortable(4).g = b + 88 / 255
    colortable(4).b = b + 88 / 255
    
    colortable(5).r = b + 225 / 255
    colortable(5).g = b + 144 / 255
    colortable(5).b = b + 88 / 255
    
    colortable(7).r = b + 225 / 255
    colortable(7).g = b + 199 / 255
    colortable(7).b = b + 88 / 255
    
    colortable(3).r = b + 144 / 255
    colortable(3).g = b + 225 / 255
    colortable(3).b = b + 88 / 255
    
    colortable(8).r = b + 88 / 255
    colortable(8).g = b + 225 / 255
    colortable(8).b = b + 199 / 255
    
    colortable(0).r = b + 88 / 255
    colortable(0).g = b + 199 / 255
    colortable(0).b = b + 225 / 255
    
    colortable(6).r = b + 88 / 255
    colortable(6).g = b + 144 / 255
    colortable(6).b = b + 225 / 255
    
    colortable(2).r = b + 167 / 255
    colortable(2).g = b + 113 / 255
    colortable(2).b = b + 250 / 255
    
    colortable(1).r = b + 225 / 255
    colortable(1).g = b + 88 / 255
    colortable(1).b = b + 199 / 255
    
    colortable(9).r = b + 255 / 255
    colortable(9).g = b + 88 / 255
    colortable(9).b = b + 144 / 255
End Sub


'centers window
Public Sub Center(ByRef win As Form)
    win.top = (Screen.height - win.height) \ 2
    win.Left = (Screen.width - win.width) \ 2
End Sub


'assert
Public Sub ASSERT(ByVal b As Boolean, Optional ByVal str As String)
    If Not b Then
        MsgBox "ASSERT: " & str, vbCritical
        End
    End If
End Sub


'returns size of file in a safe way
Public Function GetFileSize(ByVal fname As String) As Long
    On Error Resume Next
    GetFileSize = FileLen(fname)
    On Error GoTo 0
End Function


'returns file size as human readable string
Public Function FormatFileSize(ByVal size As Long) As String
Dim str As String
    If size > 1024 Then
        size = size / 1024
        If size > 1024 Then
            size = size / 1024
            str = Round(size, 2) & " MB"
        Else
            str = Round(size, 2) & " KB"
        End If
    Else
        str = size & " bytes"
    End If
    FormatFileSize = str
End Function


'cleans up file path
Public Function CleanFilePath(ByVal path As String) As String
    path = Replace(path, "\", "/")
    path = Replace(path, "//", "/")
    CleanFilePath = LCase(path)
End Function


'color3f constructor
Public Function color3f(ByVal r As Single, ByVal g As Single, ByVal b As Single) As color3f
    color3f.r = r
    color3f.g = g
    color3f.b = b
End Function


'color4f constructor
Public Function color4f(ByVal r As Single, ByVal g As Single, ByVal b As Single, ByVal a As Single) As color4f
    color4f.r = r
    color4f.g = g
    color4f.b = b
    color4f.a = a
End Function
