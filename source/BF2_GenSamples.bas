Attribute VB_Name = "BF2_SamplesGen"
Option Explicit

Public Const DEGENERATEFACEANGLE = 0.001 'angle in degrees

Private mapsizex As Long
Private mapsizey As Long
Private samplepadding As Long
Private useEdgeMargin As Boolean
Public SAMP_IgnoreTriErrors As Boolean


'writes samples for file
Public Sub WriteSamplesFile(ByVal lod As Long, ByVal w As Long, ByVal h As Long, ByVal chan As Long, _
                            ByVal padding As Long, ByVal xEdgeMargin As Boolean)
    On Error GoTo 0
    With vmesh
        'mesh must be loaded
        If Not .loadok Then
            MsgBox "No mesh loaded, or loaded with error!", vbExclamation
            Exit Sub
        End If
        
        'warning if more than one geom
        If .geomnum <> 1 Then
            MsgBox "Cannot generate samples for meshes with more than one geom", vbExclamation
            Exit Sub
        End If
        
        '4 geoms should be enough for everyone
        If .geomnum > 4 Then
            MsgBox "More than 4 LODs not supported, sorry!", vbExclamation
            Exit Sub
        End If
        
        'check size
        If Not IsPowerOfTwo(w) Or Not IsPowerOfTwo(h) Then
            MsgBox "Lightmap dimensions must be power of two!", vbExclamation
            Exit Sub
        End If
        
        'generate filename
        Dim fname As String
        fname = GetFilePath(vmesh.filename) & GetNameFromFileName(vmesh.filename)
        If lod = 0 Then
            fname = fname & ".samples"
        Else
            fname = fname & ".samp_0" & lod
        End If
        
        'check if file exists if we don't want to overwrite
        If FileExist(fname) Then
            If MsgBox(fname & " already exists." & vbLf & "Do you want to replace it?", vbExclamation Or vbYesNoCancel) <> vbYes Then
                Exit Sub
            End If
        End If
        
        'write samples
        mapsizex = w
        mapsizey = h
        samplepadding = padding
        useEdgeMargin = xEdgeMargin
        WriteLodSamples .geom(0).lod(lod), chan, fname
        
    End With
End Sub


'returns array vertex index
Private Function GetArrayVertIndex(ByRef mat As bf2mat, ByVal f As Long, ByVal fv As Long) As Long
    GetArrayVertIndex = (mat.vstart + vmesh.Index(mat.istart + (f * 3) + fv)) * (vmesh.vertstride / 4)
End Function


'replaces normal if a component is fucked
Private Function FixNormal(ByRef n As float3, ByRef rep As float3) As float3
    If IsNaN(n.X) Or IsNaN(n.Y) Or IsNaN(n.z) Then
        n = rep
    End If
End Function


'writes samples file for LOD
Public Function WriteLodSamples(ByRef lod As bf2lod, ByVal chan As Long, ByVal filename As String) As Boolean
    On Error GoTo errorhandler
    
    'MsgBox filename
    
    With lod
        Dim m As Long
        Dim f As Long
        
        'allocate samples
        Dim samplenum As Long
        Dim sample() As smp_sample
        samplenum = mapsizex * mapsizey
        ReDim sample(0 To samplenum - 1)
        
        'generate samples
        GenBF2Samples lod, sample(), samplenum, chan
        GenSamplePadding sample(), samplenum
        
        'generate face data
        Dim facenum As Long
        Dim face() As smp_face
        
        'compute number of faces
        For m = 0 To .matnum - 1
            facenum = facenum + (.mat(m).inum / 3)
        Next m
        
        'allocate faces
        ReDim face(0 To facenum - 1)
        
        'fill face array
        Dim fi As Long
        Dim stride As Long
        stride = vmesh.vertstride / 4
        For m = 0 To .matnum - 1
            With .mat(m)
                For f = 0 To (.inum / 3) - 1
                    
                    'Dim v1 As Long
                    'Dim v2 As Long
                    'Dim v3 As Long
                    'v1 = .vstart + vmesh.index(.istart + (j * 3) + 0)
                    'v2 = .vstart + vmesh.index(.istart + (j * 3) + 1)
                    'v3 = .vstart + vmesh.index(.istart + (j * 3) + 2)
                    
                    Dim f1 As Long
                    Dim f2 As Long
                    Dim f3 As Long
                    'f1 = v1 * stride
                    'f2 = v2 * stride
                    'f3 = v3 * stride
                    f1 = GetArrayVertIndex(lod.mat(m), f, 0)
                    f2 = GetArrayVertIndex(lod.mat(m), f, 1)
                    f3 = GetArrayVertIndex(lod.mat(m), f, 2)
                    
                    'vertices
                    face(fi).v1.X = vmesh.vert(f1 + 0)
                    face(fi).v1.Y = vmesh.vert(f1 + 1)
                    face(fi).v1.z = vmesh.vert(f1 + 2)
                    
                    face(fi).v2.X = vmesh.vert(f2 + 0)
                    face(fi).v2.Y = vmesh.vert(f2 + 1)
                    face(fi).v2.z = vmesh.vert(f2 + 2)
                    
                    face(fi).v3.X = vmesh.vert(f3 + 0)
                    face(fi).v3.Y = vmesh.vert(f3 + 1)
                    face(fi).v3.z = vmesh.vert(f3 + 2)
                    
                    'normals
                    face(fi).n1.X = vmesh.vert(f1 + 3 + 0)
                    face(fi).n1.Y = vmesh.vert(f1 + 3 + 1)
                    face(fi).n1.z = vmesh.vert(f1 + 3 + 2)
                    
                    face(fi).n2.X = vmesh.vert(f2 + 3 + 0)
                    face(fi).n2.Y = vmesh.vert(f2 + 3 + 1)
                    face(fi).n2.z = vmesh.vert(f2 + 3 + 2)
                    
                    face(fi).n3.X = vmesh.vert(f3 + 3 + 0)
                    face(fi).n3.Y = vmesh.vert(f3 + 3 + 1)
                    face(fi).n3.z = vmesh.vert(f3 + 3 + 2)
                    
                    fi = fi + 1
                Next f
            End With
        Next m
        
        '--- write file -----------------------------------------------------------------
        
        'delete old file
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
        
        '--- sample data --------------------
        
        'dimensions (2x 4 bytes)
        Put #ff, , mapsizex
        Put #ff, , mapsizey
        
        'samples
        Put #ff, , sample()
        
        '--- mesh data ----------------------
        
        'number of faces (4 bytes)
        Put #ff, , facenum
        
        'faces (72 byte stride)
        Put #ff, , face()
                
    End With
    
    'debug: output samples face index map
    'Dim i As Long
    'Dim foo() As Byte
    'ReDim foo(0 To samplenum - 1)
    'For i = 0 To samplenum - 1
    '    foo(i) = Clamp(sample(i).face Mod 255, 0, 255)
    'Next i
    'WriteTGA App.path & "\_sample_debug_id_.tga", Sqr(samplenum), Sqr(samplenum), 8, foo()
    
    'cleanup
    Erase sample()
        
    'close file
    Close ff
        
    'success
    WriteLodSamples = True
    Exit Function
errorhandler:
    MsgBox "WriteLodSamples" & vbLf & err.description, vbCritical
End Function


'generates sample array
Private Sub GenBF2Samples(ByRef lod As bf2lod, _
                          ByRef sample() As smp_sample, ByVal samplenum As Long, ByVal uvchan As Long)

    On Error GoTo errhandler

Dim vi1 As Long 'face vertex 1 index
Dim vi2 As Long 'face vertex 2 index
Dim vi3 As Long 'face vertex 3 index
Dim v1 As float3
Dim v2 As float3
Dim v3 As float3
Dim t1 As float2
Dim t2 As float2
Dim t3 As float2
Dim n1 As float3
Dim n2 As float3
Dim n3 As float3

Dim minx As Long 'triangle bounding rectangle in pixel space
Dim miny As Long 'triangle bounding rectangle in pixel space
Dim maxx As Long 'triangle bounding rectangle in pixel space
Dim maxy As Long 'triangle bounding rectangle in pixel space

Dim X As Long
Dim Y As Long
Dim i As Long
Dim m As Long
Dim f As Long

Dim p As float2
Dim sx As Single 'texel size in UV space
Dim sy As Single 'texel size in UV space
Dim ox As Single 'texel offset to pixel center in UV space
Dim oy As Single 'texel offset to pixel center in UV space

Dim edgemargin As Single 'margin added to rasterize pixels around triangle edges

Dim errloc As Long
    errloc = 0
    
    'compute texel scale
    sx = 1 / mapsizex
    sy = 1 / mapsizey
    
    'compute texel center offset
    ox = sx / 2
    oy = sy / 2
    
    'edge margin
    If useEdgeMargin Then
        edgemargin = max(sx, sy)
    Else
        edgemargin = 0
    End If
    
    'sample flag
    ' if 0 then sample is not rasterized
    ' if 1 then sample is inside triangle interior, sample cannot be overwritten
    ' if 2 then sample is inside triangle edge margin, may be replaced (by interior triangle sample only)
    Dim sampleflag() As Byte
    ReDim sampleflag(0 To samplenum - 1)
    
    'clear samples
    For i = 0 To samplenum - 1
        sample(i).face = -1
        sampleflag(i) = 0
    Next i
    
    'compute vertex attribute offsets
    Dim normoffset As Long
    Dim texoffset As Long
    normoffset = 3
    texoffset = 7 + (uvchan * 2)
    'normoffset = BF2MeshGetNormOffset
    'texoffset = BF2MeshGetTexcOffset(4)
        
    errloc = 1
    
    'render faces
    Dim badface As Boolean
    Dim badfacecount As Long
    Dim fi As Long
    For m = 0 To lod.matnum - 1
        With lod.mat(m)
            
            'MsgBox "mat: " & m
            
            'render faces to sample map
            For f = 0 To (.inum / 3) - 1
                badface = False
                
                errloc = 2
                
                'get vertex indices
                vi1 = GetArrayVertIndex(lod.mat(m), f, 0)
                vi2 = GetArrayVertIndex(lod.mat(m), f, 1)
                vi3 = GetArrayVertIndex(lod.mat(m), f, 2)
                
               ' ASSERT vi1 >= 0, "vertex index out of range"
               ' ASSERT vi2 >= 0, "vertex index out of range"
               ' ASSERT vi3 >= 0, "vertex index out of range"
                'If vi1 >= vmesh.vertnum Then MsgBox vi1
                'If vi2 >= vmesh.vertnum Then MsgBox vi2
                'If vi3 >= vmesh.vertnum Then MsgBox vi3
               ' ASSERT vi1 < vmesh.vertnum, "vertex index out of range"
               ' ASSERT vi2 < vmesh.vertnum, "vertex index out of range"
               ' ASSERT vi3 < vmesh.vertnum, "vertex index out of range"
                
                'get UVs
                t1.X = vmesh.vert(vi1 + texoffset + 0)
                t1.Y = vmesh.vert(vi1 + texoffset + 1)
                
                t2.X = vmesh.vert(vi2 + texoffset + 0)
                t2.Y = vmesh.vert(vi2 + texoffset + 1)
                
                t3.X = vmesh.vert(vi3 + texoffset + 0)
                t3.Y = vmesh.vert(vi3 + texoffset + 1)
                
                errloc = 3
                
                'skip triangle if it is very thin
                If Not SAMP_IgnoreTriErrors Then
                    Dim a1 As Single
                    Dim a2 As Single
                    Dim a3 As Single
                    a1 = AngleBetweenVectors(SubFloat3(v1, v2), SubFloat3(v1, v3))
                    a2 = AngleBetweenVectors(SubFloat3(v2, v1), SubFloat3(v2, v3))
                    a3 = AngleBetweenVectors(SubFloat3(v3, v2), SubFloat3(v3, v1))
                    If a1 < DEGENERATEFACEANGLE Then badface = True
                    If a2 < DEGENERATEFACEANGLE Then badface = True
                    If a3 < DEGENERATEFACEANGLE Then badface = True
                End If
                
                errloc = 4
                                
                'ASSERT t1.x >= 0 And t1.x <= 1, "t1.x out of range"
                'ASSERT t1.y >= 0 And t1.y <= 1, "t1.y out of range"
                
                'ASSERT t2.x >= 0 And t2.x <= 1, "t2.x out of range"
                'ASSERT t2.y >= 0 And t2.y <= 1, "t2.y out of range"
                
                'ASSERT t3.x >= 0 And t3.x <= 1, "t3.x out of range"
                'ASSERT t3.y >= 0 And t3.y <= 1, "t3.y out of range"
                
                If Not badface Then
                    
                    errloc = 41
                    
                    'compute triangle rect bounds
                    Dim fminx As Single
                    Dim fminy As Single
                    Dim fmaxx As Single
                    Dim fmaxy As Single
                    fminx = min(min(t1.X, t2.X), t3.X) * (mapsizex - 1)
                    fminy = min(min(t1.Y, t2.Y), t3.Y) * (mapsizey - 1)
                    fmaxx = max(max(t1.X, t2.X), t3.X) * (mapsizex - 1)
                    fmaxy = max(max(t1.Y, t2.Y), t3.Y) * (mapsizey - 1)
                    
                    errloc = 42
                    
                    minx = CLng(Round(fminx, 3)) 'can't fix overflow here, must be compiler bug or something
                    miny = CLng(Round(fminy, 3))
                    maxx = CLng(Round(fmaxx, 3))
                    maxy = CLng(Round(fmaxy, 3))
                    
                    errloc = 5
                    
                    'take in account triangle edge margin
                    If edgemargin > 0 Then
                        minx = minx - 1
                        miny = miny - 1
                        maxx = maxx + 1
                        maxy = maxy + 1
                    End If
                    
                    'clamp bounds to map size
                    minx = Clamp(minx, 0, mapsizex - 1)
                    miny = Clamp(miny, 0, mapsizey - 1)
                    maxx = Clamp(maxx, 0, mapsizex - 1)
                    maxy = Clamp(maxy, 0, mapsizey - 1)
                    
                    'check if bounds are greater than 0
                    If maxx - minx = 0 Then badface = True
                    If maxy - miny = 0 Then badface = True
                    
                End If
                
                errloc = 6
                
                'filter out bad triangles
                If Not badface Then
                    
                    'loop through rect pixels
                    For X = minx To maxx
                        For Y = miny To maxy
                            
                            'compute pixel index
                            i = CPos(X, Y)
                            
                            'only samples that are not rasterized yet or edge margin samples
                            If sampleflag(i) <> 1 Then
                            
                                'compute UV position (texel center)
                                p.X = (X * sx) + ox
                                p.Y = (Y * sy) + oy
                                
                                errloc = 7
                                
                                'triangle test
                                Dim tritest As Long
                                tritest = InsideTriangle(t1, t2, t3, p, edgemargin)
                                
                                errloc = 8
                                
                                Dim rasterize As Boolean
                                If tritest = 0 Then
                                    'triangle test failed, don't rasterize anything
                                    rasterize = False
                                ElseIf tritest = 1 Then
                                    'interior sample, always rasterize these (can overwrite edge margin samples)
                                    rasterize = True
                                ElseIf tritest = 2 Then
                                    If sampleflag(i) = 0 Then
                                        rasterize = True 'no sample yet, rasterize edge margin
                                        
                                        'modify point so it isn't outside triangle
                                        p = ClosestPointOnTriangle(t1, t2, t3, p)
                                        
                                    Else
                                        rasterize = False 'already have edge margin sample, keep it
                                    End If
                                End If
                                
                                errloc = 9
                                
                                'rasterize
                                If rasterize Then
                                    
                                    'face vertex 1
                                    v1.X = vmesh.vert(vi1 + 0)
                                    v1.Y = vmesh.vert(vi1 + 1)
                                    v1.z = vmesh.vert(vi1 + 2)
                                    
                                    'face vertex 2
                                    v2.X = vmesh.vert(vi2 + 0)
                                    v2.Y = vmesh.vert(vi2 + 1)
                                    v2.z = vmesh.vert(vi2 + 2)
                                    
                                    'face vertex 3
                                    v3.X = vmesh.vert(vi3 + 0)
                                    v3.Y = vmesh.vert(vi3 + 1)
                                    v3.z = vmesh.vert(vi3 + 2)
                                    
                                    'face normal 1
                                    n1.X = vmesh.vert(vi1 + normoffset + 0)
                                    n1.Y = vmesh.vert(vi1 + normoffset + 1)
                                    n1.z = vmesh.vert(vi1 + normoffset + 2)
                                    
                                    'face normal 2
                                    n2.X = vmesh.vert(vi2 + normoffset + 0)
                                    n2.Y = vmesh.vert(vi2 + normoffset + 1)
                                    n2.z = vmesh.vert(vi2 + normoffset + 2)
                                    
                                    'face normal 3
                                    n3.X = vmesh.vert(vi3 + normoffset + 0)
                                    n3.Y = vmesh.vert(vi3 + normoffset + 1)
                                    n3.z = vmesh.vert(vi3 + normoffset + 2)
                                    
                                    errloc = 10
                                    
                                    'fix normal if singularity
                                    Dim facenorm As float3
                                    facenorm = GenNormal(v1, v2, v3)
                                    FixNormal n1, facenorm
                                    FixNormal n2, facenorm
                                    FixNormal n3, facenorm
                                    
                                    errloc = 11
                                    
                                    'set samples
                                    sample(i).pos = TexelToPoint(v1, v2, v3, t1, t2, t3, p)
                                    sample(i).dir = Normalize(TexelToPoint(n1, n2, n3, t1, t2, t3, p))
                                    sample(i).face = fi
                                    sampleflag(i) = tritest
                                    
                                End If
                                
                            End If
                            
                        Next Y
                    Next X
                Else
                    badfacecount = badfacecount + 1
                End If
                badface = False
                
                'increment face index
                fi = fi + 1
            Next f
            
        End With
    Next m
    
    'warning
    If badfacecount > 0 Then
        MsgBox "Warning, bad " & badfacecount & " faces encountered!", vbExclamation
    End If
    
    'Dim foo() As Byte
    'ReDim foo(0 To samplenum - 1)
    'For i = 0 To samplenum - 1
    '    If sampleflag(i) = 0 Then foo(i) = 0
    '    If sampleflag(i) = 1 Then foo(i) = 255
    '    If sampleflag(i) = 2 Then foo(i) = 127
    'Next i
    'WriteTGA App.path & "\_sample_debug_.tga", Sqr(samplenum), Sqr(samplenum), 8, foo()
    
    Exit Sub
errhandler:
    MsgBox "GenBF2Samples Error Code: " & errloc & vbLf & err.description, vbCritical
End Sub


'computes pixel index from coordinates
Private Function CPos(ByVal X As Long, ByVal Y As Long) As Long
    CPos = X + (Y * mapsizex)
End Function


'applies padding to samples
Private Sub GenSamplePadding(ByRef sample() As smp_sample, ByVal samplenum As Long)
Dim i As Long  'padding iterator
Dim X As Long  'column iterator
Dim Y As Long  'row iterator
Dim j As Long  'found sample index
Dim s As Long  'temp sample index
Dim cs As Long 'current sample index
    
    On Error GoTo errh
    
    Dim tmp() As Long
    ReDim tmp(0 To samplenum - 1)
    
    For i = 1 To samplepadding
        
        'clear
        For j = 0 To samplenum - 1
            tmp(j) = -1
        Next j
        
        For X = 0 To mapsizex - 1
            For Y = 0 To mapsizey - 1
                
                cs = CPos(X, Y)
                If sample(cs).face = -1 Then
                    
                    j = -1
                    
                    'north
                    If j < 0 Then
                        If Y - 1 > 0 Then
                            s = CPos(X, Y - 1)
                            If sample(s).face > -1 Then j = s
                        End If
                    End If
                    
                    'south
                    If j < 0 Then
                        If Y + 1 < mapsizey - 1 Then
                            s = CPos(X, Y + 1)
                            If sample(s).face > -1 Then j = s
                        End If
                    End If
                    
                    'west
                    If j < 0 Then
                        If X - 1 > 0 Then
                            s = CPos(X - 1, Y)
                            If sample(s).face > -1 Then j = s
                        End If
                    End If
                    
                    'east
                    If j < 0 Then
                        If X + 1 < mapsizex - 1 Then
                            s = CPos(X + 1, Y)
                            If sample(s).face > -1 Then j = s
                        End If
                    End If
                    
                    'copy sample
                    If j > -1 Then
                        tmp(cs) = j
                        'sample(cs).pos = sample(j).pos
                        'sample(cs).dir = sample(j).dir
                        'sample(cs).face = sample(j).face
                    End If
                    
                End If
            Next Y
        Next X
        
        'sync
        For j = 0 To samplenum - 1
            If tmp(j) > -1 Then
                sample(j) = sample(tmp(j))
            End If
        Next j
        
    Next i
    
    'clean up
    Erase tmp()
    
    Exit Sub
errh:
    MsgBox err.description
End Sub
