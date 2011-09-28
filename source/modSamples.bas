Attribute VB_Name = "BF2_Samples"
Option Explicit

'sample sample structure
Public Type smp_sample      '28 bytes
    pos As float3           '12 bytes
    dir As float3           '12 bytes
    face As Long            '8 bytes
End Type

'sample face structure
Public Type smp_face       '72 bytes
    v1 As float3
    n1 As float3
    
    v2 As float3
    n2 As float3
    
    v3 As float3
    n3 As float3
End Type

'sample file structure
Public Type smp_file
    
    'header
    fourcc As String * 4        '4 bytes
    width As Long               '4 bytes
    height As Long              '4 bytes
    
    'data
    data() As smp_sample
    
    'faces
    facenum As Long
    face() As smp_face
    
    'internal
    datanum As Long             'number of samples
    filename As String
    loaded As Boolean
End Type


'sample data
Public bf2samples(0 To 4) As smp_file


'loads samples data from file
Public Function LoadSamples(ByRef smp As smp_file, ByRef filename As String) As Boolean
    On Error GoTo errorhandler
    
    'unload data
    UnloadSamples smp
    
    'check if file exists
    If Not FileExist(filename) Then
        Exit Function
    End If
    
    'open file
    Dim ff As Integer
    ff = FreeFile
    Open filename For Binary Access Read Lock Write As #ff
    
    With smp
        'copy filename
        .filename = filename
        
        'clear flag
        .loaded = False
        
        'read header
        Get #ff, , .fourcc
        Get #ff, , .width
        Get #ff, , .height
        
        If .fourcc <> "SMP2" Then
            MsgBox "File " & Chr(34) & filename & Chr(34) & " is not a supported sample file."
            Close ff
            Exit Function
        End If
        
        'determine number of samples
        .datanum = .width * .height
        Echo ">>> number of samples: " & .datanum
        Echo ">>>  size: " & .width & "x" & .height
        
        'read data (28 byte stride)
        ReDim .data(0 To .datanum - 1)
        Get #ff, , .data()
        
        Echo ">>> end of samples block at " & loc(ff)
        
        'facenum 4 bytes
        Get #ff, , .facenum
        Echo ">>> number of faces: " & .facenum
        
        'faces (72 byte stride)
        ReDim .face(0 To .facenum - 1)
        Get #ff, , .face()
        
        'check overrun
        Echo ">>> done " & loc(ff)
        Echo ">>> size " & LOF(ff)
        If loc(ff) <> LOF(ff) Then
            Echo ">>> file size does not add up!"
        End If
        
        'set flag
        .loaded = True
        
    End With
    
    'close file
    Close ff
    
    'success
    LoadSamples = True
    Exit Function
errorhandler:
    MsgBox "LoadSamples" & vbLf & err.description, vbCritical
End Function


'aligns samples with face normals
Public Sub FlattenSamples(ByRef smp As smp_file)
    On Error GoTo errhandler
    With smp
        If Not .loaded Then Exit Sub
        
        'fix samples
        Dim i As Long
        For i = 0 To .datanum - 1
            If .data(i).face > -1 Then
                
                'get face vertices
                Dim v1 As float3
                Dim v2 As float3
                Dim v3 As float3
                v1 = .face(.data(i).face).v1
                v2 = .face(.data(i).face).v2
                v3 = .face(.data(i).face).v3
                
                'compute face normal
                Dim n As float3
                n = GenNormal(v3, v2, v1)
                
                'set corrected direction vector
                .data(i).dir = n
                
            End If
        Next i
        
    End With
    
    'error handler
    Exit Sub
errhandler:
    MsgBox "FixSamples" & vbLf & err.description, vbCritical
End Sub



'fix bad samples
Public Sub FixSamples(ByRef smp As smp_file)
    On Error GoTo errhandler
    With smp
        If Not .loaded Then Exit Sub
        
        Dim valid As Long
        Dim degeneratefaces As Long
        Dim fix_renormalize As Long
        Dim fix_recompute As Long
        Dim fix_disable As Long
        
        Dim i As Long
        
        'disable samples with NaN position
        For i = 0 To .datanum - 1
            If .data(i).face > -1 Then
        
                Dim badface As Boolean
                badface = False
                
                If IsNaN(.data(i).pos.x) Then badface = True
                If IsNaN(.data(i).pos.y) Then badface = True
                If IsNaN(.data(i).pos.z) Then badface = True
        
                If badface Then
                    .data(i).face = -1
                    .data(i).pos = float3(0, 0, 0)
                    .data(i).dir = float3(0, 0, 0)
                    fix_disable = fix_disable + 1
                End If
            End If
        Next i
        
        'check for degenerate triangles
        Dim degenerate() As Byte
        ReDim degenerate(0 To .facenum - 1)
        For i = 0 To .facenum - 1
            
            'assume face is ok
            degenerate(i) = 0
            
            'check if face corner angles
            Dim a1 As Single
            Dim a2 As Single
            Dim a3 As Single
            a1 = AngleBetweenVectors(SubFloat3(.face(i).v1, .face(i).v2), SubFloat3(.face(i).v1, .face(i).v3))
            a2 = AngleBetweenVectors(SubFloat3(.face(i).v2, .face(i).v1), SubFloat3(.face(i).v2, .face(i).v3))
            a3 = AngleBetweenVectors(SubFloat3(.face(i).v3, .face(i).v2), SubFloat3(.face(i).v3, .face(i).v1))
            
            If a1 < DEGENERATEFACEANGLE Then
                degenerate(i) = 1
                degeneratefaces = degeneratefaces + 1
            End If
            If a2 < DEGENERATEFACEANGLE Then
                degenerate(i) = 1
                degeneratefaces = degeneratefaces + 1
            End If
            If a3 < DEGENERATEFACEANGLE Then
                degenerate(i) = 1
                degeneratefaces = degeneratefaces + 1
            End If
        Next i
        
        'disable samples on degenerate faces
        'For i = 0 To .datanum - 1
        '    If .data(i).face > -1 Then
        '        If degenerate(.data(i).face) = 1 Then
        '            .data(i).face = -1
        '            .data(i).pos = float3(0, 0, 0)
        '            .data(i).dir = float3(0, 0, 0)
        '            fix_disable = fix_disable + 1
        '        End If
        '    End If
        'Next i
               
        'fix samples
        For i = 0 To .datanum - 1
            If .data(i).face > -1 Then
                
                'count valid samples
                valid = valid + 1
                
                'get direction vector
                Dim v As float3
                v.x = .data(i).dir.x
                v.y = .data(i).dir.y
                v.z = .data(i).dir.z
                
                'check if direction vector components are NaN
                Dim nan As Boolean
                nan = False
                If IsNaN(v.x) Then nan = True
                If IsNaN(v.y) Then nan = True
                If IsNaN(v.z) Then nan = True
                
                'determine what to do with sample
                Dim recompute As Boolean
                recompute = False
                If nan Then
                    
                    'check if face is degenerate or not
                    If degenerate(.data(i).face) = 1 Then
                        
                        'nothing we can do, disable this sample
                        .data(i).face = -1
                        .data(i).pos = float3(0, 0, 0)
                        .data(i).dir = float3(0, 0, 0)
                        
                        fix_disable = fix_disable + 1
                    Else
                        
                        'otherwise recompute
                        recompute = True
                    End If
                Else
                    
                    'compute direction vector magnitude
                    Dim mag As Single
                    mag = Magnitude(v)
                    
                    'fix non-normalized vectors
                    If mag < 0.9 Then
                        
                        'if it is not so bad, re-normalize
                        If mag > 0.1 Then
                            v = Normalize(v)
                            
                            fix_renormalize = fix_renormalize + 1
                        Else
                            
                            'nothing we can do, recompute
                            recompute = True
                            
                        End If
                        
                    End If
                    
                End If
                
                'recompute sample direction vector
                If recompute Then
                    
                    'get face vertices
                    Dim v1 As float3
                    Dim v2 As float3
                    Dim v3 As float3
                    v1 = .face(.data(i).face).v1
                    v2 = .face(.data(i).face).v2
                    v3 = .face(.data(i).face).v3
                    
                    'compute face normal
                    v = GenNormal(v3, v2, v1)
                    
                    'Note: In theory we could compute a smooth normal by taking the
                    '      vertex normals but these are likely to be NaN as well.
                    
                    fix_recompute = fix_recompute + 1
                    
                End If
                
                'set corrected direction vector
                .data(i).dir = v
                
            End If
        Next i
        
        'display statistics
        
        'total number of fixes
        Dim fix_total As Long
        fix_total = fix_renormalize + fix_recompute + fix_disable
              
        MsgBox "Fixed " & fix_total & " samples of " & valid & ":" & vbLf & _
               "* " & degeneratefaces & " degenerate faces" & vbLf & _
               "* " & fix_renormalize & " renormalized" & vbLf & _
               "* " & fix_recompute & " recomputed" & vbLf & _
               "* " & fix_disable & " disabled", vbInformation
        
    End With
    
    'error handler
    Exit Sub
errhandler:
    MsgBox "FixSamples" & vbLf & err.description, vbCritical
End Sub


'write samples
Public Function WriteSamples(ByRef smp As smp_file, ByVal filename As String) As Boolean
    On Error GoTo errhandler
    
    If Not smp.loaded Then Exit Function
    
    'write file
    If FileExist(filename) Then
        Kill filename
    End If
    
    'create file
    Dim ff As Integer
    ff = FreeFile
    Open filename For Binary As #ff
    
    'fourcc (4 bytest)
    Dim fourcc As String * 4
    fourcc = "SMP2"
    Put #ff, , fourcc
    
    With smp
    
        '--- sample data --------------------
        
        'dimensions (2x 4 bytes)
        Put #ff, , .width
        Put #ff, , .height
        
        'samples
        Put #ff, , .data()
        
        '--- mesh data ----------------------
        
        'number of faces (4 bytes)
        Put #ff, , .facenum
        
        'faces (72 byte stride)
        Put #ff, , .face()
    
    End With
    
    'close file
    Close #ff
    
    'success
    WriteSamples = True
    Exit Function
errhandler:
    MsgBox "WriteSamples" & vbLf & err.description, vbCritical
End Function


'draws samples data
Public Sub DrawSamples()
Dim i As Long
Dim vec As float3
    
    Dim si As Long
    If bf2samples(0).loaded Then
        si = 0
    Else
        If Not view_samples Then Exit Sub
        
        si = sellod + 1
        If si < 1 Then Exit Sub
        If si > 4 Then Exit Sub
    End If
    
    With bf2samples(si)
        If Not .loaded Then Exit Sub
        
        Dim drawfaces As Boolean
        If si = 0 Then
            drawfaces = view_poly
        Else
            drawfaces = Not view_poly
        End If
        
        'draw faces
        If drawfaces Then
            If 1 = 2 Then
                glColor3f 0, 0, 0
                glEnable GL_POLYGON_OFFSET_FILL
                glPolygonOffset -3, -3
                'glColorMask False, False, False, False
                glBegin GL_TRIANGLES
                For i = 0 To .facenum - 1
                    glVertex3fv .face(i).v3.x
                    glVertex3fv .face(i).v2.x
                    glVertex3fv .face(i).v1.x
                Next i
                glEnd
                'glColorMask True, True, True, True
                glDisable GL_POLYGON_OFFSET_FILL
            End If
            
            glPolygonOffset 1, 1
            If view_lighting Then
                glColor3f 0.3, 0.5, 0.3
                glEnable GL_POLYGON_OFFSET_FILL
                glEnable GL_LIGHTING
                glBegin GL_TRIANGLES
                For i = 0 To .facenum - 1
                    glNormal3fv .face(i).n1.x
                    glVertex3fv .face(i).v1.x
                    
                    glNormal3fv .face(i).n2.x
                    glVertex3fv .face(i).v2.x
                    
                    glNormal3fv .face(i).n3.x
                    glVertex3fv .face(i).v3.x
                Next i
                glEnd
                glDisable GL_LIGHTING
            Else
                glBegin GL_TRIANGLES
                For i = 0 To .facenum - 1
                    glColor3f .face(i).n1.x * 0.5 + 0.5, .face(i).n1.y * 0.5 + 0.5, .face(i).n1.z * 0.5 + 0.5
                    glVertex3fv .face(i).v1.x
                    
                    glColor3f .face(i).n2.x * 0.5 + 0.5, .face(i).n2.y * 0.5 + 0.5, .face(i).n2.z * 0.5 + 0.5
                    glVertex3fv .face(i).v2.x
                    
                    glColor3f .face(i).n3.x * 0.5 + 0.5, .face(i).n3.y * 0.5 + 0.5, .face(i).n3.z * 0.5 + 0.5
                    glVertex3fv .face(i).v3.x
                Next i
                glEnd
            End If
            glDisable GL_POLYGON_OFFSET_FILL
            
        End If
        
        If view_samples Then
            
            'draw lines
            StartAALine 1.3
            glBegin GL_LINES
            glColor3f 1, 0.5, 0
            For i = 0 To .datanum - 1
                If .data(i).face > -1 Then
                    
                    vec.x = .data(i).pos.x + (.data(i).dir.x * 0.05)
                    vec.y = .data(i).pos.y + (.data(i).dir.y * 0.05)
                    vec.z = .data(i).pos.z + (.data(i).dir.z * 0.05)
                    glVertex3fv .data(i).pos.x
                    glVertex3fv vec.x
                    
                End If
            Next i
            glEnd
            EndAALine
            
            'draw points
            StartAAPoint 4
            glColor3f 1, 1, 0
            glBegin GL_POINTS
            For i = 0 To .datanum - 1
                If .data(i).face > -1 Then
                    glVertex3fv .data(i).pos.x
                End If
            Next i
            glEnd
            EndAAPoint
            
        End If
        
        'draw bad samples
        StartAAPoint 8
        glColor3f 1, 0, 0
        glBegin GL_POINTS
        For i = 0 To .datanum - 1
            If .data(i).face > -1 Then
                Dim nan As Boolean
                nan = False
                If IsNaN(.data(i).dir.x) Then nan = True
                If IsNaN(.data(i).dir.y) Then nan = True
                If IsNaN(.data(i).dir.z) Then nan = True
                If nan Then glVertex3fv .data(i).pos.x
            End If
        Next i
        glEnd
        EndAAPoint
        
    End With
End Sub


'draws 2d bitmap
Public Sub DrawSamples2d(ByRef pic As PictureBox)
Dim i As Long
Dim x As Long
Dim y As Long
Dim si As Long
Dim r As Single
Dim g As Single
Dim b As Single
    
    If bf2samples(0).loaded Then
        si = 0
    Else
        si = sellod + 1
        If si < 1 Then Exit Sub
        If si > 4 Then Exit Sub
    End If
    
    pic.Cls
    With bf2samples(si)
        For x = 0 To .width - 1
            For y = 0 To .height - 1
                
                i = x + (y * .width)
                
                If .data(i).face > -1 Then
                    r = (.data(i).dir.x + 1) * 127.5
                    g = (.data(i).dir.y + 1) * 127.5
                    b = (.data(i).dir.z + 1) * 127.5
                End If
                
                pic.PSet (4 + x, 4 + y), RGB(r, g, b)
                
            Next y
        Next x
    End With
    
End Sub


'unloads samples file
Public Function UnloadSamples(ByRef smp As smp_file)
    With smp
        .loaded = False
        .fourcc = ""
        .width = 0
        .height = 0
        .datanum = 0
        Erase .data()
    End With
End Function

'--- samples files -------------------------------------------------


'loads all samples associated with mesh file
Public Sub LoadSamplesFiles(ByRef meshfilename As String)

    Dim fname As String
    'fname = GetNameFromFileName(meshfilename)
    fname = Replace(meshfilename, ".staticmesh", "")
    
    LoadSamples bf2samples(1), fname & ".samples"
    LoadSamples bf2samples(2), fname & ".samp_01"
    LoadSamples bf2samples(3), fname & ".samp_02"
    LoadSamples bf2samples(4), fname & ".samp_03"
    
End Sub


'unloads all samples
Public Sub UnloadSamplesFiles()
Dim i As Long
    For i = 1 To 4
        UnloadSamples bf2samples(i)
    Next i
End Sub


'--- treeview -------------------------------------------------------

'fills treeview hierarchy
Public Sub FillTreeSamples(ByRef tree As MSComctlLib.TreeView)
Dim i As Long
Dim n As MSComctlLib.node
Dim rootname As String
Dim addedroot As Boolean
    
    'lods
    For i = 0 To 4
        With bf2samples(i)
            If .loaded Then
                
                'If Not addedroot Then
                '    addedroot = True
                '
                '    rootname = "samples_root"
                '    Set n = tree.Nodes.Add(, , rootname, "Samples", "geom")
                '    n.Expanded = True
                'End If
                
                Dim sampname As String
                sampname = "samp" & i
                Set n = tree.Nodes.Add(, tvwChild, sampname, GetFileName(.filename), "file")
                If i = 0 Then n.Expanded = True
                
                tree.Nodes.Add sampname, tvwChild, sampname & "|trinum", "Polygons: " & .facenum, "trinum"
                tree.Nodes.Add sampname, tvwChild, sampname & "|sampnum", "Samples: " & .datanum, "prop"
                tree.Nodes.Add sampname, tvwChild, sampname & "|width", "Width: " & .width, "prop"
                tree.Nodes.Add sampname, tvwChild, sampname & "|height", "Height: " & .height, "prop"
                
            End If
        End With
    Next i
    
End Sub
