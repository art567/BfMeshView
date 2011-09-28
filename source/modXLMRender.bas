Attribute VB_Name = "a_LMRender"
Option Explicit

'settings
Public lmoutput As String
Public lmoutputalpha As Boolean
Public lmwidth As Long
Public lmheight As Long
Public lmwarnoverwrite As Boolean
Public lmshowoutput As Boolean
Public lmoutputnormals As Boolean

Public lmres As Long
Public lmframesize As Long
Public lmfov As Single
Public lmnear As Single
Public lmfar As Single
Public lmpasses As Long
Public lmpadding As Long
Public lmtwosided As Boolean
Public lmhemisphere As Boolean

Public lmaaa As Boolean
Public lmaaathres As Long
Public lmaaapasses As Long 'todo: add to config, GUI etc.

Public lmaccel As Boolean
Public lmaccelthres As Long

Public lmfalloff As Boolean
Public lmfalloffstart As Single
Public lmfalloffend As Single

'states
Public lmrender As Boolean
Public lmabort As Boolean
Public lmpause As Boolean

Private Type sample_type
    pos As float3
    vec As float3
    flag As Boolean         'render flag
    alpha As Byte           'alpha mask
    tempc As Byte           'temp color
End Type


'private stuff
Private dlist As GLuint
Private mapsizex As Long
Private mapsizey As Long
Private mapdata() As Byte           'map buffer
Private bucket() As Byte            'frame buffer readback array
Private sample() As sample_type
Private samplenum As Long
Private faceorder() As Long
Private hemivert(0 To 6 - 1) As float3
Private usedepth As Boolean
Private var_index() As Integer
Private var_indexnum As Long
Private drawmode As Long

'renders lighting map
Public Sub RenderLighting()
    
    lmrender = True
    
    drawmode = 2
    
    'map size
    mapsizex = lmwidth / lmres
    mapsizey = lmheight / lmres
    samplenum = mapsizex * mapsizey
    
    'show output window
    If lmshowoutput Then
        If Not frmOutput.Visible Then frmOutput.Show
        frmOutput.SetSize mapsizex, mapsizey
    End If
    
    'allocate
    ReDim bucket(0 To (lmframesize * lmframesize) - 1)
    ReDim mapdata(0 To samplenum - 1)
    ReDim sample(0 To samplenum - 1)
    
    'generate sample map
    GenSampleMap
    
'If 1 = 2 Then
    
    'generate optimized mesh
    LMXCreateMesh
    
    'setup viewport
    glClearColor 1, 1, 1, 0
    glViewport 0, 0, lmframesize, lmframesize
    glColor3f 0, 0, 0
    
    'setup projection
    glMatrixMode GL_PROJECTION
    glLoadIdentity
    gluPerspective lmfov, 1, lmnear, lmfar
    glMatrixMode GL_MODELVIEW
    
    'prepare some things
    glColorMask True, False, False, False
    glDisable GL_TEXTURE_2D
    glDisable GL_LIGHTING
    glPolygonMode GL_FRONT_AND_BACK, GL_FILL
    glFrontFace GL_CCW
    
    'two sided
    If lmtwosided Then
        glDisable GL_CULL_FACE
    Else
        glEnable GL_CULL_FACE
    End If
    
    'falloff fog
    If lmfalloff Then
        Dim fc(4) As Single
        fc(0) = 1
        fc(1) = 1
        fc(2) = 1
        fc(3) = 1
        glFogfv GL_FOG_COLOR, fc(0)
        glFogi GL_FOG_MODE, GL_LINEAR
        glFogf GL_FOG_START, lmfalloffstart
        glFogf GL_FOG_END, lmfalloffend
        glEnable GL_FOG
        usedepth = True
    Else
        usedepth = False
    End If
    
    'hemisphere
    If lmhemisphere Then
        Dim s As Single
        s = (lmnear + lmfar) / 2
        hemivert(0).y = -s
        hemivert(1).x = -s:  hemivert(1).z = s
        hemivert(2).x = -s:  hemivert(2).z = -s
        hemivert(3).x = s:   hemivert(3).z = -s
        hemivert(4).x = s:   hemivert(4).z = s
        hemivert(5).x = -s:  hemivert(5).z = s
    End If
    
    'timer
    Dim t1 As Long
    Dim t2 As Long
    Dim sec As Single
    t1 = timeGetTime
    
    'render passes
    Dim prog As Long
    Dim oldprog As Long
    Dim pass As Long
    frmRender.pgbPass.max = 100
    frmRender.pgbTotal.max = lmpasses
    For pass = 1 To lmpasses
        
        'progress bar update
        frmRender.pgbTotal.Value = pass
        DoEvents
        
        'depth testing
        If pass > 1 Then usedepth = True
        If usedepth Then
            glEnable GL_DEPTH_TEST
        Else
            glDisable GL_DEPTH_TEST
        End If
        
        'render samples
        Dim i As Long
        Dim j As Long
        Dim lastdraw As Long
        lastdraw = 0
        For i = 0 To samplenum - 1
            
            If sample(i).flag Then
                mapdata(i) = RenderSample(sample(i))
                sample(i).alpha = 255
            Else
                mapdata(i) = 0
                sample(i).alpha = 0
            End If
            
            'update progress
            prog = (i / samplenum) * 100
            If prog > oldprog Then
                
                t2 = timeGetTime
                sec = Round((t2 - t1) / 1000, 1)
                frmRender.labStats.Caption = "Samples: " & i & "/" & samplenum & " Time taken: " & sec & " sec"
                
                frmRender.pgbPass.Value = prog
                
                'draw to output window
                If lmshowoutput Then
                    Dim x As Long
                    Dim y As Long
                    For j = lastdraw To i
                        x = j Mod mapsizex
                        y = Fix(j / mapsizex)
                        frmOutput.SetPixel x, y, mapdata(j)
                    Next j
                    lastdraw = i
                End If
                
                DoEvents
                oldprog = prog
                
                If lmpause Then
                    Do While lmpause
                        DoEvents
                    Loop
                End If
                
                If lmabort Then Exit For
            End If
            
        Next i
        
    Next pass
    
    'reset some things
    If lmfalloff Then glDisable GL_FOG
    If Not usedepth Then glEnable GL_DEPTH_TEST
    glColorMask True, True, True, True
    
    'apply padding
    GenPadding
    
'End If
    
    'save output
    WriteTGA lmoutput, mapsizex, mapsizey, 8, mapdata()
    
    'save normal map
    
    
    'clean up
    Erase bucket()
    Erase sample()
    Erase mapdata()
    LMXDestroyMesh
    
    'finish
    lmrender = False
End Sub


'swaps the face index of two faceorders
Private Sub SwapFaceOrder(ByVal f1 As Long, ByVal f2 As Long)
Dim t As Long
    t = faceorder(f1)
    faceorder(f1) = faceorder(f2)
    faceorder(f2) = t
End Sub


'generates face order array
Private Function GenFaceOrder(ByVal g As Long)
Dim f As Long
Dim i As Long

Dim v1 As float3
Dim v2 As float3
Dim v3 As float3
Dim vv1 As float3
Dim vv2 As float3
Dim vv3 As float3

Dim t1 As float2
Dim t2 As float2
Dim t3 As float2
Dim tt1 As float2
Dim tt2 As float2
Dim tt3 As float2

Dim n As float3
Dim p As float4
Dim d As Single
    
    With myobj
        
        'reset priority
        For i = 0 To .group(g).facenum - 1
            faceorder(i) = i
        Next i
        
        'compute priority for all triangles
        For f = 0 To .group(g).facenum - 1
            
            'get vertices
            v1 = .vert(.group(g).face(f).v1)
            v2 = .vert(.group(g).face(f).v2)
            v3 = .vert(.group(g).face(f).v3)
            
            'get texcoords
            t1 = .texc(.group(g).face(f).t1)
            t2 = .texc(.group(g).face(f).t2)
            t3 = .texc(.group(g).face(f).t3)
            
            'calculate triangle normal
            n = GenNormal(v1, v2, v3)
            
            'create plane
            p.x = -n.x
            p.y = -n.y
            p.z = -n.z
            p.w = DotProduct(n, v1)
            
            'check against other faces
            For i = 0 To .group(g).facenum - 1
                
                'don't check against self
                If i <> f Then
                    
                    'get texcoords
                    tt1 = .texc(.group(g).face(i).t1)
                    tt2 = .texc(.group(g).face(i).t2)
                    tt3 = .texc(.group(g).face(i).t3)
                    
                    'check for overlap
                    If TriTriOverlapTest(t1, t2, t3, tt1, tt2, tt3) Then
                        
                        'get vertex
                        vv1 = .vert(.group(g).face(i).v1)
                        vv2 = .vert(.group(g).face(i).v2)
                        vv3 = .vert(.group(g).face(i).v3)
                        
                        'plane test
                        If PlaneTest(p, vv1) < 0 Then
                            If PlaneTest(p, vv2) < 0 Then
                                If PlaneTest(p, vv3) < 0 Then
                                    SwapFaceOrder i, f
                                End If
                            End If
                        End If
                        
                    End If
                    
                End If
                
            Next i
            
        Next f
        
    End With
    
End Function


'generates sample array
Private Sub GenSampleMap()
Dim g As Long
Dim f As Long
Dim v1 As float3
Dim v2 As float3
Dim v3 As float3
Dim t1 As float2
Dim t2 As float2
Dim t3 As float2
Dim n1 As float3
Dim n2 As float3
Dim n3 As float3

Dim minx As Single
Dim miny As Single
Dim maxx As Single
Dim maxy As Single

Dim x As Long
Dim y As Long
Dim i As Long
Dim j As Long

Dim p As float2
Dim sx As Single
Dim sy As Single
Dim ox As Single
Dim oy As Single
    
    'compute scale
    sx = 1 / mapsizex
    sy = 1 / mapsizey
    
    'compute offset
    ox = sx / 2
    oy = sy / 2
    
    'clear samples
    For i = 0 To samplenum - 1
        sample(i).flag = False
    Next i
    
    'render faces
    With myobj
        For g = 0 To .groupnum - 1
            
            'determine face ordering to solve overlap
            ReDim faceorder(0 To .group(g).facenum - 1)
            GenFaceOrder g
            
            'render faces to sample map
            For j = 0 To .group(g).facenum - 1
                f = faceorder((.group(g).facenum - 1) - j)
                'Echo "face " & (f + 1)
                
                'get UVs
                t1 = .texc(.group(g).face(f).t1)
                t2 = .texc(.group(g).face(f).t2)
                t3 = .texc(.group(g).face(f).t3)
                
                'compute triangle rect bounds
                minx = min(min(t1.x, t2.x), t3.x) * (mapsizex - 1)
                miny = min(min(t1.y, t2.y), t3.y) * (mapsizey - 1)
                maxx = max(max(t1.x, t2.x), t3.x) * (mapsizex - 1)
                maxy = max(max(t1.y, t2.y), t3.y) * (mapsizey - 1)
                minx = Clamp(minx, 0, mapsizex - 1)
                miny = Clamp(miny, 0, mapsizey - 1)
                maxx = Clamp(maxx, 0, mapsizex - 1)
                maxy = Clamp(maxy, 0, mapsizey - 1)
                
                'loop through rect pixels
                For x = minx To maxx
                    For y = miny To maxy
                    
                        'compute pixel index
                        i = x + (y * mapsizex)
                        
                        If sample(i).flag = False Then
                        
                            'compute UV position (texel center)
                            p.x = (x * sx) + ox
                            p.y = (y * sy) + oy
                            
                            'triangle test
                            If InsideTriangle(t1, t2, t3, p, 0) > 0 Then
                                
                                'get vertices
                                v1 = .vert(.group(g).face(f).v1)
                                v2 = .vert(.group(g).face(f).v2)
                                v3 = .vert(.group(g).face(f).v3)
                                
                                'get normals
                                n1 = .norm(.group(g).face(f).n1)
                                n2 = .norm(.group(g).face(f).n2)
                                n3 = .norm(.group(g).face(f).n3)
                                
                                'set samples
                                sample(i).pos = TexelToPoint(v1, v2, v3, t1, t2, t3, p)
                                sample(i).vec = Normalize(TexelToPoint(n1, n2, n3, t1, t2, t3, p))
                                sample(i).flag = True
                                
                                ''''hack
                                'mapdata(i) = 55 + ((faceorder(f) / (.group(g).facenum - 1)) * 200)
                                ''''hack
                                
                            End If
                            
                        End If
                        
                    Next y
                Next x
                
            Next j
            
            'clean up
            Erase faceorder()
            
        Next g
    End With
    
End Sub


'applies padding to texture
Private Sub GenPadding()
Dim i As Long
Dim x As Long
Dim y As Long
Dim n As Long
Dim c As Long
Dim a As Long
    'On Error GoTo errh
    
    For i = 1 To lmpadding
        For x = 0 To mapsizex - 1
            For y = 0 To mapsizey - 1
                If sample(CPos(x, y)).alpha = 0 Then
                    
                    c = 0
                    n = 0
                    
                    'north
                    If y > 0 Then
                        If sample(CPos(x, y - 1)).alpha = 255 Then
                            c = c + mapdata(CPos(x, y - 1))
                            n = n + 1
                        End If
                    End If
                    
                    'south
                    If y < mapsizey - 1 Then
                        If sample(CPos(x, y + 1)).alpha = 255 Then
                            c = c + mapdata(CPos(x, y + 1))
                            n = n + 1
                        End If
                    End If
                    
                    'west
                    If x > 0 Then
                        If sample(CPos(x - 1, y)).alpha = 255 Then
                            c = c + mapdata(CPos(x - 1, y))
                            n = n + 1
                        End If
                    End If
                    
                    'east
                    If x < mapsizex - 1 Then
                        If sample(CPos(x + 1, y)).alpha = 255 Then
                            c = c + mapdata(CPos(x + 1, y))
                            n = n + 1
                        End If
                    End If
                    
                    'set pixel
                    If n > 0 Then
                        mapdata(CPos(x, y)) = (c / n)
                        sample(CPos(x, y)).alpha = 127
                    End If
                    
                End If
            Next y
        Next x
        
        'update alpha
        For x = 0 To mapsizex - 1
            For y = 0 To mapsizey - 1
                If sample(CPos(x, y)).alpha = 127 Then
                    sample(CPos(x, y)).alpha = 255
                End If
            Next y
        Next x
        
    Next i
    
    Exit Sub
errh:
    MsgBox Err.Description
End Sub

Private Sub MergeSamples(ByRef a As sample_type, ByRef b As sample_type, dst As sample_type)
    dst.pos.x = (a.pos.x + b.pos.x) * 0.5
    dst.pos.y = (a.pos.x + b.pos.y) * 0.5
    dst.pos.z = (a.pos.x + b.pos.y) * 0.5
    dst.vec.x = (a.vec.x + b.vec.x) * 0.5
    dst.vec.y = (a.vec.x + b.vec.y) * 0.5
    dst.vec.z = (a.vec.x + b.vec.y) * 0.5
    dst.vec = Normalize(dst.vec)
End Sub

'adaptive anti-aliasing pass
Private Sub RenderAAA()
Dim x As Long
Dim y As Long
Dim i As Long
Dim j As Long
Dim c As Byte
Dim ct As Byte
Dim diff As Long
Dim samp As sample_type
    
    For x = 0 To mapsizex - 1
        For y = 0 To mapsizey - 1
            i = CPos(x, y)
            If sample(i).flag Then
                
                'get pixel
                c = mapdata(i)
                
                'set temp color
                sample(i).tempc = c
                
                'north
                If y > 0 Then
                    j = CPos(x, y - 1)
                    If sample(j).flag Then
                        ct = mapdata(CPos(x, y - 1))
                        diff = Abs(c - ct)
                        If diff > lmaaathres Then
                            'If Distance(sample(i).pos, sample(j).pos) < rad Then
                                MergeSamples sample(i), sample(j), samp
                                c = RenderSample(samp)
                                '''
                            'End If
                        End If
                    End If
                End If
                
                'south
                If y < mapsizey - 1 Then
                    j = CPos(x, y + 1)
                    If sample(j).flag Then
                        ct = mapdata(CPos(x, y + 1))
                        diff = Abs(c - ct)
                        If diff > lmaaathres Then
                            
                        End If
                    End If
                End If
                
                'west
                If x > 0 Then
                    j = CPos(x - 1, y)
                    If sample(j).flag Then
                        ct = mapdata(CPos(x - 1, y))
                        diff = Abs(c - ct)
                        If diff > lmaaathres Then
                            
                        End If
                    End If
                End If
                
                'east
                If x < mapsizex - 1 Then
                    j = CPos(x + 1, y)
                    If sample(j).flag Then
                        ct = mapdata(CPos(x + 1, y))
                        diff = Abs(c - ct)
                        If diff > lmaaathres Then
                            
                        End If
                    End If
                End If
                
            End If
        Next y
    Next x
    
    'update refresh
    For i = 0 To samplenum - 1
        mapdata(i) = sample(i).tempc
    Next i
    
End Sub


'renders a single pixel
Private Function RenderSample(ByRef s As sample_type) As Byte
Dim up As float3
    
    'clear buffers
    If usedepth Then
        glClear GL_COLOR_BUFFER_BIT Or GL_DEPTH_BUFFER_BIT
    Else
        glClear GL_COLOR_BUFFER_BIT
    End If
    
    'determine up vector
    up.x = 0
    up.y = 1
    up.z = 0
    If s.vec.y > 0.9 Then
        up.x = 1
        up.y = 0
        up.z = 0
    End If
    
    'setup view
    glLoadIdentity
    gluLookAt s.pos.x, s.pos.y, s.pos.z, _
             (s.pos.x + s.vec.x), (s.pos.y + s.vec.y), (s.pos.z + s.vec.z), _
              up.x, up.y, up.z
    
    'draw hemisphere
    If lmhemisphere Then
        glColor3f 0, 0, 0
        If usedepth Then glDepthMask False
        glPushMatrix
            glTranslatef s.pos.x, s.pos.y, s.pos.z
            DrawHemisphere
        glPopMatrix
        If usedepth Then glDepthMask True
    End If
    
    'draw mesh
    LMXDrawMesh
    
    'get result
    glFinish
    RenderSample = GetAverage()
    
    'SwapBuffers frmMain.picMain.hDC 'temp
End Function


'returns average framebuffer intensity
Private Function GetAverage() As Byte
Dim i As Long
Dim s As Long
Dim a As Long
    s = lmframesize * lmframesize
    glReadPixels 0, 0, lmframesize, lmframesize, GL_RED, GL_UNSIGNED_BYTE, ByVal VarPtr(bucket(0))
    For i = 0 To s - 1
        a = a + bucket(i)
    Next i
    GetAverage = a / s
End Function


'computes pixel index from coordinates
Private Function CPos(ByVal x As Long, ByVal y As Long) As Long
    CPos = x + (mapsizex * y)
End Function


'draws hemisphere
Public Sub DrawHemisphere()
    glVertexPointer 3, GL_FLOAT, 0, ByVal VarPtr(hemivert(0).x)
    glEnableClientState GL_VERTEX_ARRAY
    glDrawArrays GL_TRIANGLE_FAN, 0, 6
    glDisableClientState GL_VERTEX_ARRAY
End Sub


'--- scene drawing ----------------------------------------------------------------------------------------------------


'builds optimized mesh
Private Sub LMXCreateMesh()
Dim i As Long
Dim j As Long
    
    'display list path
    If drawmode = 1 Then
        dlist = glGenLists(1)
        glNewList dlist, GL_COMPILE
            glBegin GL_TRIANGLES
            With myobj
                For i = 0 To .groupnum - 1
                    With .group(i)
                        For j = 0 To .facenum - 1
                            glVertex3fv myobj.vert(.face(j).v1).x
                            glVertex3fv myobj.vert(.face(j).v2).x
                            glVertex3fv myobj.vert(.face(j).v3).x
                        Next j
                    End With
                Next i
            End With
            glEnd
        glEndList
    End If
    
    'VAR path
    If drawmode = 2 Then
        Dim g As Long
        Dim f As Long
        With myobj
            
            'compute indexnum
            var_indexnum = 0
            For g = 0 To .groupnum - 1
                var_indexnum = var_indexnum + (.group(g).facenum * 3)
            Next g
            
            'fill index array
            ReDim var_index(0 To var_indexnum - 1)
            i = 0
            For g = 0 To .groupnum - 1
                For f = 0 To .group(g).facenum - 1
                    var_index(i + 0) = .group(g).face(f).v1
                    var_index(i + 1) = .group(g).face(f).v2
                    var_index(i + 2) = .group(g).face(f).v3
                    i = i + 3
                Next f
            Next g
            
        End With
    End If
    
    'VBO path
    If drawmode = 3 Then
        '
    End If
    
End Sub


'renders optmized mesh
Private Sub LMXDrawMesh()
    
    'display list path
    If drawmode = 1 Then
        glCallList dlist
    End If
    
    'VAR path
    If drawmode = 2 Then
        glEnableClientState GL_VERTEX_ARRAY
        glVertexPointer 3, GL_FLOAT, 0, ByVal VarPtr(myobj.vert(0))
        glDrawElements GL_TRIANGLES, var_indexnum, GL_UNSIGNED_SHORT, ByVal VarPtr(var_index(0))
        glDisableClientState GL_VERTEX_ARRAY
    End If
    
    'VBO path
    If drawmode = 3 Then
        '
    End If
    
End Sub


'destroys optimized mesh
Private Sub LMXDestroyMesh()
    
    'display list path
    If drawmode = 1 Then
        glDeleteLists dlist, 1
    End If
    
    'VAR path
    If drawmode = 2 Then
        Erase var_index()
        var_indexnum = 0
    End If
    
    'VBO path
    If drawmode = 3 Then
        '
    End If
    
End Sub

