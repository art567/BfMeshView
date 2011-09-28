Attribute VB_Name = "BF2_Renormalize"
Option Explicit

Public Sub Renormalize()
Dim i As Long
Dim j As Long

Dim stride As Long

Dim vi As Long
Dim vj As Long

Dim sharedvertnum As Long
Dim sharedvert() As Long
    
    With vmesh
        
        'pre-allocate
        sharedvertnum = 0
        ReDim sharedvert(0 To .vertnum - 1)
        
        stride = .vertstride / 4
        
        'add all vertices
        For i = 0 To .vertnum - 1
            vi = i * stride
            
            'check if already added
            For j = 0 To sharedvertnum - 1
                vj = sharedvert(sharedvertnum - 1)
                
                'compare vertex
                If .vert(vi + 0) = .vert(vj + 0) Then
                If .vert(vi + 1) = .vert(vj + 1) Then
                If .vert(vi + 2) = .vert(vj + 2) Then
                    
                    'compare normal
                    If .vert(vi + 3 + 0) = .vert(vj + 3 + 0) Then
                    If .vert(vi + 3 + 1) = .vert(vj + 3 + 1) Then
                    If .vert(vi + 3 + 2) = .vert(vj + 3 + 2) Then
                        
                        sharedvertnum = sharedvertnum + 1
                        sharedvert(sharedvertnum - 1) = vi
                        
                    End If
                    End If
                    End If
                    
                End If
                End If
                End If
                
            Next j
        Next i
        
    End With
End Sub
