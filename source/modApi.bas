Attribute VB_Name = "a_Api"
Option Explicit

Public Declare Function DrawRect Lib "gdi32" Alias "Rectangle" (ByVal hdc As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long

Public Declare Function timeGetTime Lib "winmm.dll" () As Long

Public Declare Function ShellExecute Lib "shell32" Alias "ShellExecuteA" (ByVal hWnd As Long, _
                                                                          ByVal lpOperation As String, _
                                                                          ByVal lpFile As String, _
                                                                          ByVal lpParameters As String, _
                                                                          ByVal lpDirectory As String, _
                                                                          ByVal nShowCmd As Long) As Long

Private Declare Function CompareMemory Lib "Ntdll.dll" Alias "RtlCompareMemory" (buffA As Any, buffB As Any, ByVal Length As Long) As Long
Public Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (ByVal dst As Long, ByVal src As Long, ByVal num As Long)

Public Declare Function LockWindowUpdate Lib "user32" (ByVal hWnd As Long) As Long

Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, _
                                                    ByVal X As Long, ByVal Y As Long, _
                                                    ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Declare Sub DragAcceptFiles Lib "shell32.dll" (ByVal hWnd As Long, ByVal fAccept As Long)

Private Declare Function DragQueryFile Lib "shell32.dll" Alias "DragQueryFileA" (ByVal HDROP As Long, _
                                                                                 ByVal UINT As Long, _
                                                                                 ByVal lpStr As String, _
                                                                                 ByVal ch As Long) As Long

Private Declare Sub DragFinish Lib "shell32.dll" (ByVal HDROP As Long)

Private Declare Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, _
                                                                                  ByVal hWnd As Long, _
                                                                                  ByVal uMsg As Long, _
                                                                                  ByVal wParam As Long, _
                                                                                  ByVal lParam As Long) As Long

Private Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hWnd As Long, _
                                                                            ByVal wMsg As Long, _
                                                                            ByVal wParam As Long, _
                                                                            ByVal lParam As Long) As Long

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, _
                                                                            ByVal nIndex As Long) As Long

Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hWnd As Long, _
                                                                                ByVal nIndex As Long, _
                                                                                ByVal dwNewLong As Long) As Long

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type MINMAXINFO
    ptReserved As POINTAPI
    ptMaxSize As POINTAPI
    ptMaxPosition As POINTAPI
    ptMinTrackSize As POINTAPI
    ptMaxTrackSize As POINTAPI
End Type

'Public Const MK_CONTROL = &H8
'Public Const MK_LBUTTON = &H1
'Public Const MK_RBUTTON = &H2
'Public Const MK_MBUTTON = &H10
'Public Const MK_SHIFT = &H4
Private Const GWL_WNDPROC = -4
Private Const GWL_STYLE = -16
Private Const WM_DROPFILES = &H233
Private Const WM_MOUSEWHEEL = &H20A
Private Const WM_GETMINMAXINFO = &H24
Private Const TVS_NOTOOLTIPS = &H80
Private Const SWP_NOMOVE = 2
Private Const SWP_NOSIZE = 1
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const NANBASE = &H7FF80000


Private hook_hwnd As Long
Private hook_prevwndproc As Long


'hooks window
Public Sub Hook(ByRef hWnd As Long)
    
    'wndproc
    hook_hwnd = hWnd
    hook_prevwndproc = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf WndProc)
    
    'drag and drop
    DragAcceptFiles hook_hwnd, True
    
End Sub


'window callback
Private Function WndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim MouseKeys As Long
Dim value As Long
Dim X As Long
Dim Y As Long
    
    Select Case uMsg
    Case WM_MOUSEWHEEL:
        MouseKeys = wParam And 65535
        value = (wParam / 65536) / 120
        X = lParam And 65535
        Y = lParam / 65536
        frmMain.MouseWheel MouseKeys, value, X, Y
    Case WM_DROPFILES
        Dim str As String * 512
        DragQueryFile wParam, 0, str, 512
        frmMain.DropFile safeStr(str)
        DragFinish wParam
    Case WM_GETMINMAXINFO
        Dim mm As MINMAXINFO
        CopyMem VarPtr(mm), lParam, LenB(mm)
        mm.ptMinTrackSize.X = 400
        mm.ptMinTrackSize.Y = 300
        CopyMem lParam, VarPtr(mm), LenB(mm)
        WndProc = DefWindowProc(hWnd, uMsg, wParam, lParam)
        Exit Function
    End Select
    
    WndProc = CallWindowProc(hook_prevwndproc, hWnd, uMsg, wParam, lParam)
End Function


'unhooks window
Public Sub UnHook()
    
    'drag and drop
    DragAcceptFiles hook_hwnd, False
    
    'wndproc
    SetWindowLong hook_hwnd, GWL_WNDPROC, hook_prevwndproc
    
End Sub


'make the window topmost
Public Sub SetTopMostWindow(ByVal hWnd As Long, ByVal top As Boolean)
    If top Then
        Call SetWindowPos(hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
    Else
        Call SetWindowPos(hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
    End If
End Sub


'disables annoying treeview tooltips
Public Sub DisableTreeViewToolTips(ByVal hWnd As Long)
    Dim style As Long
    style = GetWindowLong(hWnd, GWL_STYLE)
    SetWindowLong hWnd, GWL_STYLE, style Or TVS_NOTOOLTIPS
End Sub


'returns NAN
Private Function GetNaN____xx() As Single
    'CopyMem VarPtr(GetNaN) + 4, VarPtr(NANBASE), 4
    CopyMem VarPtr(GetNaN), VarPtr(NANBASE), 4
End Function
Public Function GetNaN() As Single
    On Error Resume Next
    GetNaN = 0 / 0
    On Error GoTo 0
End Function


'returns whether a number is NaN or not
Public Function IsNaN(ByRef v As Single) As Boolean
    'VB is shit, touch a NaN and we're dead!
    'so we convert it to string and do a string compare
    
    Dim s As String
    s = CStr(v)
    
    If s = "1.#IND" Then
        IsNaN = True
        Exit Function
    End If
    
    If s = "-1.#IND" Then
        IsNaN = True
        Exit Function
    End If
    
    If s = "1.#INF" Then
        IsNaN = True
        Exit Function
    End If
    
    If s = "-1.#INF" Then
        IsNaN = True
        Exit Function
    End If
    
    If s = "1.#QNAN" Then
        IsNaN = True
        Exit Function
    End If
    
    If s = "-1.#QNAN" Then
        IsNaN = True
        Exit Function
    End If
    
    IsNaN = False
End Function


'checks float3 for NaNs
Public Function IsNaN3f(ByRef v As float3) As Boolean
    If IsNaN(v.X) Then
        IsNaN3f = True
        Exit Function
    End If
    If IsNaN(v.Y) Then
        IsNaN3f = True
        Exit Function
    End If
    If IsNaN(v.z) Then
        IsNaN3f = True
        Exit Function
    End If
    IsNaN3f = False
End Function

