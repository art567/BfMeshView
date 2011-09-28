Attribute VB_Name = "BF1942_StdMesh"
Option Explicit

'mesh file types
Private Enum stdmeshfiletype
    sm_single = 1           'single mesh?
    sm_multi = 6            'more than one mesh?
End Enum


'lod mat struct
Private Type stdmeshmat
    
    'name (? bytes)
    matname As String
    
    'data (36 bytes)
    u2 As Long          '0
    u3 As Long          '0
    u4 As Long          '0
    prim As Long        'primitive type (4==triangle, 5==triangle_strip, OpenGL enums!)
    u6 As Long          '1041, could be bit flags
    vertstride As Long  '32
    vertnum As Long
    indexnum As Long
    u7 As Long          '0/2
    
    'data (? bytes)
    vert() As Single
    Index() As Integer
    
    'internal
    texmapid As Long
    shaderid As Long 'index of RS shader
End Type


'lod struct
Private Type stdmeshlod
    matnum As Long
    mat() As stdmeshmat
    
    'internal
    polycount As Long
End Type


'various col structs
Private Type colvert '16 bytes
    v As float3
    w As Single
End Type
Private Type colface '8 bytes
    v1 As Integer
    v2 As Integer
    v3 As Integer
    matid As Byte
    flags As Byte
End Type
Private Type colq '32 bytes
    i1 As Integer 'looks like index
    i2 As Integer 'looks like index
    flags As Integer
    u1 As Integer
    u2 As Single
    u3 As Single
    u4 As Single
    u5 As Long 'looks like index
    u6 As Long 'looks like index
    u7 As Long 'looks like index
End Type


'col struct
Private Type stdmeshcol
    size As Long
    u1 As Long              '???
    u2 As Long              '???
    
    vertnum As Long         '
    vert() As colvert       '
    
    facenum As Long         '
    face() As colface       '
    
    qnum As Long
    qdata() As colq
    
    u3 As Long
    flags As Long
    ustr As String * 24
    
    znum As Long
    zdata() As Long
    
    u4 As Integer
    
    'internal
    facenorm() As float3
End Type


'file struct
Private Type stdmeshfile
    
    'header
    version As Long     '9/10
    u2 As Long          '0
    
    'bounds
    min As float3
    max As float3
    
    'unknown
    qflag As Byte
    
    'y block (unknown)
    colnum As Long
    col() As stdmeshcol
    
    'unknown
    lodnum As Long     '1/6
    u5 As Byte          '0
    
    'meshes
    lod() As stdmeshlod
    
    'c block
    cid As Long
    csize As Long
    cdata() As Byte
    
    'internal
    filename As String
    loadok As Boolean
    drawok As Boolean
End Type

Public stdmesh As stdmeshfile


'loads StandardMesh from file
Public Function LoadStdMesh(ByRef filename As String) As Boolean
    On Error GoTo errorhandler
    
Dim i As Long
Dim j As Long
    
    'open file
    Dim ff As Integer
    ff = FreeFile
    Open filename For Binary Access Read Lock Write As #ff
    
    With stdmesh
        .filename = filename
        .drawok = True
        .loadok = False
        
        '--- header -----------------------------------------------
        
        'unknown (8 bytes)
        Get #ff, , .version
        Get #ff, , .u2
        Echo "u1: " & .version
        Echo "u2: " & .u2
        Echo ""
        
        'bounds (24 bytes)
        Echo "bounds start at " & loc(ff)
        Get #ff, , .min
        Get #ff, , .max
        Echo "bounds end at " & loc(ff)
        Echo ""
        
        'unknown (1 byte)
        If .version > 9 Then
            Get #ff, , .qflag
            Echo "qflag: " & .qflag
        End If
        Echo ""
        
        '--- cols ---------------------------------------------------
        
        'colnum (4 bytes)
        Get #ff, , .colnum
        Echo "colnum: " & .colnum
        
        'cols (? bytes)
        If .colnum > 0 Then
            ReDim .col(0 To .colnum - 1)
            For i = 0 To .colnum - 1
                Echo "col " & i & " start at " & loc(ff)
                ReadStdMeshCol ff, .col(i)
                Echo "col " & i & " end at " & loc(ff)
                Echo ""
            Next i
        End If
        Echo ""
        
        '--- meshes --------------------------------------------------
        
        'meshnum (4 bytes)
        Get #ff, , .lodnum
        Echo "lodnum: " & .lodnum
        
        'meshes (? bytes)
        If .lodnum > 0 Then
            ReDim .lod(0 To .lodnum - 1)
            For i = 0 To .lodnum - 1
                Echo "lod " & i & " start at " & loc(ff)
                ReadStdMeshLod ff, .lod(i)
                Echo "lod " & i & " end at " & loc(ff)
                Echo ""
            Next i
        End If
        
        '--- unknown --------------------------------------------------
        
        Get #ff, , .cid
        Echo "cid: " & .cid
        
        '--- cdata ----------------------------------------------------
        
        Get #ff, , .csize
        Echo "csize: " & .csize
        
        If .csize > 0 Then
            ReDim .cdata(0 To .csize - 1)
            Get #ff, , .cdata()
        End If
        
        '--- end of file ----------------------------------------------
        
        Echo "done reading " & loc(ff)
        Echo "file size is " & LOF(ff)
        Echo ""
        
        .loadok = True
        .drawok = True
    End With
    
    'close file
    Close #ff
    
    LoadStdMesh = True
    Exit Function
errorhandler:
    MsgBox "LoadStdMesh" & vbLf & err.description, vbCritical
    Echo ">>> error at " & loc(ff)
    Echo ">>> file size " & LOF(ff)
    stdmesh.drawok = True 'False
End Function


'reads StandardMesh mesh chunk
Private Sub ReadStdMeshLod(ByRef ff As Integer, ByRef lod As stdmeshlod)
Dim i As Long
    With lod
        
        'matnum (4 bytes)
        Get ff, , .matnum
        Echo " matnum: " & .matnum
        
        'skip if empty
        If .matnum = 0 Then Exit Sub
                
        'read materials
        ReDim .mat(0 To .matnum - 1)
        For i = 0 To .matnum - 1
            With .mat(i)
            
                'material name (? bytes)
                .matname = ReadStdMeshString(ff)
                Echo " matname: " & .matname
                
                'unknown (16 bytes)
                Get #ff, , .u2
                Get #ff, , .u3
                Get #ff, , .u4
                Echo " u2: " & .u2
                Echo " u3: " & .u3
                Echo " u4: " & .u4
                
                'unknown (8 bytes)
                Get #ff, , .prim
                Get #ff, , .u6
                Echo " prim: " & .prim
                Echo " u6: " & .u6
                
                'mesh info (12 bytes)
                Get #ff, , .vertstride
                Get #ff, , .vertnum
                Get #ff, , .indexnum
                Echo " vertstride: " & .vertstride
                Echo " vertnum:    " & .vertnum
                Echo " indexnum:   " & .indexnum
                
                'unknown (4 bytes)
                Get #ff, , .u7
                Echo " u7: " & .u7
                
            End With
            
            .polycount = .polycount + (.mat(i).indexnum / 3)
        Next i
        
        'read mat geometry data
        For i = 0 To .matnum - 1
            With .mat(i)
                
                'read vertices
                ReDim .vert(0 To (.vertnum * (.vertstride / 4)) - 1)
                Get #ff, , .vert()
                
                'read face indices
                ReDim .Index(0 To .indexnum - 1)
                Get #ff, , .Index()
                
            End With
        Next i
        
    End With
End Sub

Private Sub ReadStdMeshCol(ByRef ff As Integer, ByRef col As stdmeshcol)
    Dim i As Long
    With col
        'block size (4 bytes)
        Get #ff, , .size
        Echo " size: " & .size
        
        Dim skip As Long
        skip = loc(ff) + .size
        
        '8 bytes
        Get #ff, , .u1          '???
        Get #ff, , .u2          'always 5?
        Echo " u1: " & .u1
        Echo " u2: " & .u2
        
        '--- vertices -------------------
        
        'vertnum (4 bytes)
        Get #ff, , .vertnum
        Echo " vertnum: " & .vertnum
        
        'vertices (vertnum * 16 bytes)
        If .vertnum > 0 Then
            Echo " verts start at " & loc(ff)
            ReDim .vert(0 To .vertnum - 1)
            Get #ff, , .vert()
            Echo " verts end at " & loc(ff)
        End If
        
        '--- faces ----------------------
        
        'facenum (4 bytes)
        Get #ff, , .facenum
        Echo " facenum: " & .facenum
        
        'faces (facenum * 8 bytes)
        If .facenum > 0 Then
            Echo " face start at " & loc(ff)
            ReDim .face(0 To .facenum - 1)
            Get #ff, , .face()
            Echo " faces end at " & loc(ff)
            
            For i = 0 To .facenum - 1
                Echo ">>> matid: " & .face(i).matid & Chr(9) & .face(i).flags
            Next i
        End If
        
        'generate normals
        If .facenum > 0 Then
            ReDim .facenorm(0 To .facenum - 1)
            For i = 0 To .facenum - 1
                .facenorm(i) = GenNormal(.vert(.face(i).v1).v, .vert(.face(i).v2).v, .vert(.face(i).v3).v)
            Next i
        End If
        
        '''temp: skip over the rest
        Seek #ff, 1 + skip
        Echo "skipped to: " & loc(ff)
        Exit Sub
        '''temp
        
        '--- qblock ---------------------
       '
       ' 'qnum (4 bytes)
       ' Get #ff, , .qnum
       ' Echo " qnum: " & .qnum
       '
       ' 'qdata (qnum * 32 bytes)
       ' If .qnum > 0 Then
       '     Echo " qdata start at " & Loc(ff)
       '     ReDim .qdata(0 To .qnum - 1)
       '     Get #ff, , .qdata()
       '     Echo " qdata end at " & Loc(ff)
       ' End If
       '
        
        Dim u1 As Long
        Dim u2 As Long
        Dim u3 As Long
        Get #ff, , u1
        Get #ff, , u2
        Get #ff, , u3
        
        'qdata (qnum * 32 bytes)
        If .facenum > 0 Then
            Echo " qdata start at " & loc(ff)
            ReDim .qdata(0 To .facenum - 1)
            Get #ff, , .qdata()
            Echo " qdata end at " & loc(ff)
        End If
        
        
        '--- ??? ------------------------
        
        Get #ff, , .u3
        Get #ff, , .flags
        Get #ff, , .ustr
        
        Echo " u3: " & .u3
        Echo " flags: " & .flags
        Echo " ustr: " & .ustr
        
        '--- zblock ---------------------
        
        'znum (4 bytes)
        Get #ff, , .znum
        Echo " znum: " & .znum
        
        'zdata (znum * 4 bytes)
        If .znum > 0 Then
            Echo " zdata start at " & loc(ff)
            ReDim .zdata(0 To .znum - 1)
            Get #ff, , .zdata()
            Echo " zdata end at " & loc(ff)
        End If
        
        '--- ??? ------------------------
        
        'u4 (2 bytes)
        Get #ff, , .u4
        Echo " u4: " & .u4
    End With
End Sub


'writes StandardMesh lod chunk
Private Sub WriteStdMeshLod(ByRef ff As Integer, ByRef lod As stdmeshlod)
Dim i As Long
    With lod
        
        'matnum (4 bytes)
        Put ff, , .matnum
        
        'skip if empty
        If .matnum = 0 Then Exit Sub
        
        'read materials
        For i = 0 To .matnum - 1
            With .mat(i)
            
                'material name (? bytes)
                WriteStdMeshString ff, .matname
                
                'unknown (16 bytes)
                Put #ff, , .u2
                Put #ff, , .u3
                Put #ff, , .u4
                
                'unknown (8 bytes)
                Put #ff, , .prim
                Put #ff, , .u6
                
                'mesh info (12 bytes)
                Put #ff, , .vertstride
                Put #ff, , .vertnum
                Put #ff, , .indexnum
                
                'unknown (4 bytes)
                Put #ff, , .u7
                
            End With
        Next i
        
        'write mat geometry data
        For i = 0 To .matnum - 1
            With .mat(i)
                
                'write vertices
                Put #ff, , .vert()
                
                'write face indices
                Put #ff, , .Index()
                
            End With
        Next i
        
    End With
End Sub


'reads string
Private Function ReadStdMeshString(ByRef ff As Integer) As String
Dim num As Long
Dim chars() As Byte
    Get #ff, , num
    
    If num = 0 Then Exit Function
    
    ReDim chars(0 To num - 1)
    Get #ff, , chars()
    
    ReadStdMeshString = SafeString(chars, num)
End Function


'reads string
Private Sub WriteStdMeshString(ByRef ff As Integer, ByRef str As String)
Dim strlen As Long
    strlen = Len(str)
    
    'write length (4 bytes)
    Put #ff, , strlen
    
    'write characters
    Dim i As Long
    For i = 1 To strlen
        Dim b As Byte
        b = Asc(Mid(str, i, 1))
        Put #ff, , b
    Next i
End Sub


'writes .standardmesh
Public Function WriteStdMesh(ByRef filename As String) As Boolean
    
    MsgBox "Currently not implemented, maybe next version again.", vbInformation
    Exit Function
    
    'NOTE: this function is broken, fix col writing!!!!!!!!!!!!!!!!!!!!!!!!!!
    
    On Error GoTo errorhandler
    
    'create file
    Dim ff As Integer
    ff = FreeFile
    Open filename For Binary As #ff
    
    Dim i As Long
    With stdmesh
        
        '--- header -----------------------------------------------
    
        'unknown (8 bytes)
        Put #ff, , .version
        Put #ff, , .u2
        
        'bounds (24 bytes)
        Put #ff, , .min
        Put #ff, , .max
        
        'unknown (1 byte)
        If .version = 10 Then
            Put #ff, , .qflag
        End If
        
        '--- y block ---------------------------------------------------
        
        'yblocknum (4 bytes)
        Put #ff, , .colnum
        
        'yblock[]
        For i = 0 To .colnum - 1
            'Put #ff, , .yblock(i).size
            'Put #ff, , .yblock(i).data()
        Next i
        
        '--- meshes --------------------------------------------------
        
        'meshnum (4 bytes)
        Put #ff, , .lodnum
        Echo " lodnum: " & .lodnum
        
        'mesh[]
        For i = 0 To .lodnum - 1
            WriteStdMeshLod ff, .lod(i)
        Next i
        
        '--- unknown --------------------------------------------------
        
        'cid (4 bytes)
        Put #ff, , .cid
        
        'csize (4 bytes)
        Put #ff, , .csize
        
        'cdata[]
        If .csize > 0 Then
            Put #ff, , .cdata()
        End If
        
        '--- end of file ----------------------------------------------
        
    End With
    
    'close file
    Close #ff
    
    'success
    WriteStdMesh = True
    Exit Function
    
    'error handler
errorhandler:
    MsgBox "WriteStdMesh" & vbLf & err.description, vbCritical
End Function



'draws StandardMesh
Public Sub DrawStdMesh()
    On Error GoTo errorhandler
    Dim i As Long
    Dim j As Long
    
    With stdmesh
        If Not .drawok Then Exit Sub
        
        If selgeom = 0 Then
            glColor3f 0.75, 0.75, 0.75
            With .lod(sellod)
                
                For j = 0 To .matnum - 1
                    With .mat(j)
                        
                        glVertexPointer 3, GL_FLOAT, .vertstride, .vert(0)
                        glNormalPointer GL_FLOAT, .vertstride, .vert(3)
                        glTexCoordPointer 2, GL_FLOAT, .vertstride, .vert(6)
                        
                        glEnableClientState GL_VERTEX_ARRAY
                        glEnableClientState GL_NORMAL_ARRAY
                        glEnableClientState GL_TEXTURE_COORD_ARRAY
                        
                        If view_poly Then
                        
                            'draw solid
                            If .shaderid And view_textures Then
                                BindStdShader .shaderid
                            Else
                                glColor3f 0.75, 0.75, 0.75
                                UnbindTexture
                                If view_lighting Then
                                    glEnable GL_LIGHTING
                                Else
                                    glDisable GL_LIGHTING
                                End If
                                If view_backfaces Then
                                    glDisable GL_CULL_FACE
                                Else
                                    glEnable GL_CULL_FACE
                                End If
                                glDisable GL_BLEND
                                glDisable GL_ALPHA_TEST
                                glDepthMask True
                            End If
                            
                            'draw polygons
                            If view_edges Or view_verts Then
                                glPolygonOffset 1, 1
                                glEnable GL_POLYGON_OFFSET_FILL
                            End If
                            Dim prim As GLenum
                            glDrawElements .prim, .indexnum, GL_UNSIGNED_SHORT, .Index(0)
                            If view_edges Or view_verts Then
                                glDisable GL_POLYGON_OFFSET_FILL
                            End If
                            
                            'reset some things
                            UnbindTexture
                            If Not view_backfaces Then
                                glEnable GL_CULL_FACE
                            End If
                            glDepthMask True
                            glDisable GL_LIGHTING
                            glDisable GL_BLEND
                            glDisable GL_ALPHA_TEST
                            
                            'draw outlines
                            If view_edges And Not view_wire Then
                                glColor4f 1, 1, 1, 0.1
                                StartAALine 1.3
                                glPolygonMode GL_FRONT_AND_BACK, GL_LINE
                                glDrawElements .prim, .indexnum, GL_UNSIGNED_SHORT, .Index(0)
                                glPolygonMode GL_FRONT_AND_BACK, GL_FILL
                                EndAALine
                            End If
                        
                         End If
                         
                         glDisable GL_TEXTURE_2D
                         glDisable GL_LIGHTING
                        
                        'draw vertices
                        If view_verts Then
                            glColor3f 1, 1, 1
                            StartAAPoint 4
                            glDrawArrays GL_POINTS, 0, .vertnum
                            EndAALine
                        End If
                        
                        glDisableClientState GL_TEXTURE_COORD_ARRAY
                        glDisableClientState GL_NORMAL_ARRAY
                        glDisableClientState GL_VERTEX_ARRAY
                        
                    End With
                Next j
                
            End With
            
            'mesh bounds
            If view_bounds Then
                StartAALine 1.3
                glColor3f 1, 1, 0
                DrawBox .min, .max
                EndAALine
            End If
            
        Else
            
            
            With .col(sellod)
                
                glColor3f 0.5, 0.75, 1
                
                If view_lighting Then glEnable GL_LIGHTING
                
                glBegin GL_TRIANGLES
                For i = 0 To .facenum - 1
                    
                    Dim cc As Integer
                    cc = Clamp(.face(i).matid Mod maxcolors, 0, maxcolors)
                    glColor4fv colortable(cc).r
                    
                    glNormal3fv .facenorm(i).X
                    glVertex3fv .vert(.face(i).v3).v.X
                    glVertex3fv .vert(.face(i).v2).v.X
                    glVertex3fv .vert(.face(i).v1).v.X
                Next i
                glEnd
                
                If view_lighting Then glDisable GL_LIGHTING
            End With
            
        End If
        
    End With
    Exit Sub
errorhandler:
    MsgBox "DrawStdMesh" & vbLf & err.description, vbCritical
    stdmesh.drawok = False
End Sub
   

'unloads StandardMesh
Public Sub UnloadStdMesh()
    With stdmesh
        .loadok = False
        .drawok = False
        .filename = ""
        
        'todo
    End With
End Sub


'fills treeview hierarchy
Public Sub FillTreeStdMesh(ByRef tree As MSComctlLib.TreeView)
    Dim i As Long
    Dim j As Long
    With stdmesh
        If Not .loadok Then Exit Sub
        
        Dim n As MSComctlLib.node
        
        'file root
        Dim rootname As String
        rootname = "stdmesh_root"
        Set n = tree.Nodes.Add(, , rootname, GetFileName(.filename), "file")
        n.Expanded = True
        n.tag = 0
        
        Set n = tree.Nodes.Add(rootname, tvwChild, "ver", "Version: " & .version, "prop")
        Set n = tree.Nodes.Add(rootname, tvwChild, "u2", "u2: " & .u2, "prop")
        
        Dim meshname As String
        Dim lodname As String
        Dim matname As String
        Dim tag As Long
        
        If .lodnum > 0 Then
            
            meshname = "mesh"
            Set n = tree.Nodes.Add(rootname, tvwChild, meshname, "Mesh", "geom")
            n.Expanded = True
            n.tag = MakeTag(0, 0, 0)
            
            'lods
            For i = 0 To .lodnum - 1
                With .lod(i)
                    
                    lodname = "lod " & i
                    tag = MakeTag(0, i, 0)
                    
                    Set n = tree.Nodes.Add(meshname, tvwChild, lodname, "Lod " & i, "lod")
                    n.tag = tag
                    
                    Set n = tree.Nodes.Add(lodname, tvwChild, lodname & "|trinum", "Polygons: " & .polycount, "trinum")
                    n.tag = tag
                    
                    'materials
                    For j = 0 To .matnum - 1
                        With .mat(j)
                            
                            matname = lodname & "|mat " & j
                            Set n = tree.Nodes.Add(lodname, tvwChild, matname, .matname, "mat")
                            n.tag = MakeTag(0, i, j)
                            
                            Set n = tree.Nodes.Add(matname, tvwChild, matname & "|vertnum", "Vertices: " & .vertnum, "prop")
                            n.tag = tag
                            
                            Set n = tree.Nodes.Add(matname, tvwChild, matname & "|indexnum", "Indices: " & .indexnum, "prop")
                            n.tag = tag
                            
                        End With
                    Next j
                End With
            Next i
        End If
        
        If .colnum > 0 Then
            
            meshname = "colmesh"
            Set n = tree.Nodes.Add(rootname, tvwChild, meshname, "Collision Mesh", "geom")
            n.Expanded = True
            n.tag = MakeTag(1, 0, 0)
            
            'cols
            For i = 0 To .colnum - 1
                With .col(i)
                    
                    lodname = "col " & i
                    tag = MakeTag(1, i, 0)
                    
                    Set n = tree.Nodes.Add(meshname, tvwChild, lodname, "Col " & i, "lod")
                    n.tag = tag
                    
                    Set n = tree.Nodes.Add(lodname, tvwChild, lodname & "|trinum", "Polygons: " & .facenum, "trinum")
                    n.tag = tag
                    
                    Set n = tree.Nodes.Add(lodname, tvwChild, lodname & "|vertnum", "Vertices: " & .vertnum, "prop")
                    n.tag = tag
                    
                End With
            Next i
        End If
        
    End With
End Sub


'returns the filename of the texture that is currently selected in treeview
Public Function GetStdMeshSelectedTextureFilename() As String
    On Error GoTo errhandler
    
    Dim texmapid As Long
    Dim shaderid As Long
    
    'select by shader node
    With stdshader
        If .loaded Then
            
            If seltex >= 1 And seltex <= .subshader_num Then
                
                texmapid = stdshader.subshader(seltex).texmapid
                
                If texmapid > 0 Then
                    GetStdMeshSelectedTextureFilename = texmap(texmapid).filename
                    Exit Function
                End If
                
            End If
            
        End If
    End With
    
    Exit Function
    
    'select by material node
    With stdmesh
        If .loadok Then
            If sellod < 0 Then Exit Function
            If sellod > .lodnum - 1 Then Exit Function
            
            If selmat < 0 Then Exit Function
            If selmat > .lod(sellod).matnum - 1 Then Exit Function
            
            shaderid = .lod(sellod).mat(selmat).shaderid
            If shaderid = 0 Then Exit Function
            
            texmapid = stdshader.subshader(.lod(sellod).mat(selmat).shaderid).texmapid
            If texmapid > 0 Then
                GetStdMeshSelectedTextureFilename = texmap(texmapid).filename
                Exit Function
            End If
        End If
    End With
    
    Exit Function
errhandler:
    MsgBox "GetStdMeshSelectedTextureFilename" & vbLf & err.description, vbCritical
End Function
