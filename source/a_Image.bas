Attribute VB_Name = "a_Image"
Option Explicit

Public Type bgr
    b As Byte
    g As Byte
    r As Byte
End Type

Public Type bgra
    b As Byte
    g As Byte
    r As Byte
    a As Byte
End Type


'applies padding to texture
Public Sub GenPadding(ByVal w As Long, ByVal h As Long, ByRef data() As bgra, ByVal padding As Long)
Dim i As Long
Dim x As Long
Dim y As Long
Dim n As Long
Dim cr As Long
Dim cg As Long
Dim cb As Long
Dim a As Long
Dim p As Long
    On Error GoTo errhandler
    
    For i = 1 To padding
        For x = 0 To w - 1
            For y = 0 To h - 1
                p = x + (y * w)
                If data(p).a = 0 Then
                    
                    cr = 0
                    cg = 0
                    cb = 0
                    n = 0
                    
                    'north
                    If y > 0 Then
                        p = x + (w * (y - 1))
                        If data(p).a = 255 Then
                            cr = cr + data(p).r
                            cg = cg + data(p).g
                            cb = cb + data(p).b
                            n = n + 1
                        End If
                    End If
                    
                    'south
                    If y < h - 1 Then
                        p = x + (w * (y + 1))
                        If data(p).a = 255 Then
                            cr = cr + data(p).r
                            cg = cg + data(p).g
                            cb = cb + data(p).b
                            n = n + 1
                        End If
                    End If
                    
                    'west
                    If x > 0 Then
                        p = (x - 1) + (y * w)
                        If data(p).a = 255 Then
                            cr = cr + data(p).r
                            cg = cg + data(p).g
                            cb = cb + data(p).b
                            n = n + 1
                        End If
                    End If
                    
                    'east
                    If x < w - 1 Then
                        p = (x + 1) + (y * w)
                        If data(p).a = 255 Then
                            cr = cr + data(p).r
                            cg = cg + data(p).g
                            cb = cb + data(p).b
                            n = n + 1
                        End If
                    End If
                    
                    'set pixel
                    If n > 0 Then
                        p = x + (y * w)
                        data(p).r = cr / n
                        data(p).g = cg / n
                        data(p).b = cb / n
                        data(p).a = 127
                    End If
                    
                End If
            Next y
        Next x
        
        'update alpha
        For x = 0 To w - 1
            For y = 0 To h - 1
                p = x + (y * w)
                If data(p).a = 127 Then
                    data(p).a = 255
                End If
            Next y
        Next x
        
    Next i
    
    Exit Sub
errhandler:
    MsgBox err.description
End Sub

