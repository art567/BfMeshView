Attribute VB_Name = "OBJ_Export"
Option Explicit


Public Function ExportMesh(ByRef filename As String) As Boolean
    On Error GoTo errorhandler
    
Dim i As Long
Dim j As Long
Dim k As Long
Dim f As Long

Dim stride As Long
Dim v1 As Long
Dim v2 As Long
Dim v3 As Long
Dim i1 As Long
Dim i2 As Long
Dim i3 As Long
    
    'open file
    Dim ff As Integer
    ff = FreeFile
    Open filename For Output As #ff
    
    With vmesh
        stride = .vertstride / 4
        
        Print #ff, "# Wavefront OBJ File"
        Print #ff, ""
        
        Print #ff, "# Vertices"
        For i = 0 To .vertnum - 1
            v1 = (i * stride) + 0 + 0
            v2 = (i * stride) + 0 + 1
            v3 = (i * stride) + 0 + 2
            Print #ff, "v " & fff(-.vert(v1)) & " " & fff(.vert(v2)) & " " & fff(.vert(v3))
        Next i
        Print #ff, ""
        
        Print #ff, "# Texture Coordinates"
        For i = 0 To .vertnum - 1
            v1 = (i * stride) + 7 + (2 * (.uvnum - 1)) + 0
            v2 = (i * stride) + 7 + (2 * (.uvnum - 1)) + 1
            Print #ff, "vt " & fff(.vert(v1)) & " " & fff(.vert(v2))
        Next i
        Print #ff, ""
        
        Print #ff, "# Vertex Normals"
        For i = 0 To .vertnum - 1
            v1 = (i * stride) + 3 + 0
            v2 = (i * stride) + 3 + 1
            v3 = (i * stride) + 3 + 2
            Print #ff, "vn " & fff(-.vert(v1)) & " " & fff(.vert(v2)) & " " & fff(.vert(v3))
        Next i
        Print #ff, ""
        
        For i = 0 To .geomnum - 1
            With .geom(i)
                Print #ff, "# Geom " & i
                Print #ff, ""
                
                For j = sellod To sellod
                'For j = 0 To .lodnum - 1
                    With .lod(j)
                        Print #ff, "# Lod " & j
                        Print #ff, "g Lod_" & j
                        'Print #ff, "s 1"
                        
                        For k = 0 To .matnum - 1
                            With .mat(k)
                                Print #ff, "usemtl Material_" & k
                                
                                For f = 0 To .inum - 1 Step 3
                                    
                                    i3 = .vstart + vmesh.Index(.istart + f + 0) + 1
                                    i2 = .vstart + vmesh.Index(.istart + f + 1) + 1
                                    i1 = .vstart + vmesh.Index(.istart + f + 2) + 1
                                    
                                    'Print #ff, "f " & i1 & " " & i2 & " " & i3
                                    
                                    Print #ff, "f " & i1 & "/" & i1 & "/" & i1 & " " & _
                                                      i2 & "/" & i2 & "/" & i2 & " " & _
                                                      i3 & "/" & i3 & "/" & i3
                                    
                                Next f
                            End With
                            
                        Next k
                        
                        Print #ff, ""
                    End With
                Next j
                
                Print #ff, ""
            End With
        Next i
        
        Print #ff, "# End of file"
    End With
    
    'close file
    Close #ff
    
    'success
    ExportMesh = True
    Exit Function
errorhandler:
    MsgBox "ExportMesh" & vbLf & err.description, vbCritical
End Function


Public Sub DrawOBJExp()

Dim i As Long
Dim j As Long
Dim k As Long
Dim f As Long

Dim stride As Long
Dim v1 As Long
Dim v2 As Long
Dim v3 As Long
Dim i1 As Long
Dim i2 As Long
Dim i3 As Long
    
    glPointSize 4
    glColor3f 0.75, 0.75, 0.75
    With vmesh
        stride = .vertstride / 4
    
        For i = 0 To .geomnum - 1
            With .geom(i)
                
                For j = 0 To .lodnum - 1
                    With .lod(j)
                        
                        For k = 0 To .matnum - 1
                            With .mat(k)
                                
                                For f = 0 To .inum - 1 Step 3
                                    
                                    i1 = .vstart + vmesh.Index(.istart + f + 0) '+ 1
                                    i2 = .vstart + vmesh.Index(.istart + f + 1) '+ 1
                                    i3 = .vstart + vmesh.Index(.istart + f + 2) '+ 1
                                    
                                    
                                    glBegin GL_TRIANGLES
                                        glVertex3fv vmesh.vert(i1 * stride)
                                        glVertex3fv vmesh.vert(i2 * stride)
                                        glVertex3fv vmesh.vert(i3 * stride)
                                        'Print #ff, "f " & i1 & " " & i2 & " " & i3
                                    glEnd
                                    
                                    'Print #ff, "f " & i1 & "/" & i1 & "/" & i1 & " " & _
                                    '                  i2 & "/" & i2 & "/" & i2 & " " & _
                                    '                  i3 & "/" & i3 & "/" & i3
                                    
                                Next f
                            End With
                            
                        Next k
                        
                    End With
                Next j
                
            End With
        Next i
    End With
End Sub

