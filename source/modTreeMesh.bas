Attribute VB_Name = "BF1942_TreeMesh"
Option Explicit


'treemesh mesh material group
Private Type tmmat
    start As Long       'index start offset
    count As Long       'index count
    texname As String   'texture filename string
    
    ''''internal
    texmapid As Long
End Type


'treemesh mesh
Private Type tmmesh
    matnum As Long
    mat() As tmmat
End Type


'collider vertex
Private Type tmcolvert '16 bytes
    x As Single
    y As Single
    z As Single
    flags As Integer
    u1 As Integer
End Type

'collider face
Private Type tmcolface '8 bytes
    v1 As Integer
    v2 As Integer
    v3 As Integer
    flags As Integer
End Type

'unknown (perhaps AABB tree)
Private Type tmhdata '32 bytes
    u1 As float3    'some coordinate (perhaps normal vector???)
    u2 As Long      'always 0?
    
    v1 As Long      'looks like vertex index
    v2 As Long      'looks like vertex index
    v3 As Long      'looks like vertex index
    
    flags As Long   'looks like collision flags (always 82?)
End Type

'collision AABB node
Private Type tmaabb
    flag As Byte
    min As float3
    max As float3
    facenum As Long
    face() As Long
    u1 As Integer
    u2 As Integer
    u3 As Integer
End Type

'tree vertex
Private Type tmvert '44 bytes
    vert As float3 '12 bytes
    norm As float3 '12 bytes
    u1 As Byte     '1 byte    'maybe float=sprite size?
    u2 As Byte     '1 byte
    u3 As Byte     '1 byte
    u4 As Byte     '1 byte
    tex As float2  '8 bytes   'UV0
    u5 As float2   '8 bytes   'UV1??
End Type


'file struct
Private Type tmfile
    
    'header
    u1 As Long          '3
    u2 As Long          '0
    u3 As Long          '8
    
    'bounds
    min As float3
    max As float3
    min2 As float3
    max2 As float3
    
    'meshes
    meshnum As Long
    mesh() As tmmesh
    
    '--- collision mesh ------------
    
    'collision data
    colflag As Long         'file has collison chunk if not null
    colu1 As Long           'always 5? could be collision chunk format version number
    
    colvertnum As Long
    colvert() As tmcolvert
    
    colfacenum As Long
    colface() As tmcolface
    
    h_u1 As Long
    h_u2 As Long
    
    'h block
    hnum As Long
    hdata() As tmhdata
    
    '--- geometry -----------------
    
    'vertices
    vertnum As Long
    vert() As tmvert
    
    'indices
    indexnum As Long
    Index() As Integer
    
    
    ''''internal
    filename As String
    loadok As Boolean
    drawok As Boolean
    colfacenorm() As float3
End Type

Public treemesh As tmfile


'loads TreeMesh from file
Public Function LoadTreeMesh(ByRef filename As String) As Boolean
    On Error GoTo errorhandler
    
Dim i As Long
Dim j As Long
    
    'open file
    Dim ff As Integer
    ff = FreeFile
    Open filename For Binary Access Read Lock Write As #ff
    
    With treemesh
        .filename = filename
        .drawok = True
        .loadok = False
        
        '--- header -----------------------------------------------
        
        Echo "header start at " & loc(ff)
        
        'unknown (8 bytes)
        Get #ff, , .u1
        Get #ff, , .u2
        Get #ff, , .u3
        Echo " u1: " & .u1
        Echo " u2: " & .u2
        Echo " u2: " & .u3
        
        'bounds (24 bytes)
        Echo " bounds start at " & loc(ff)
        Get #ff, , .min
        Get #ff, , .max
        Echo " bounds end at " & loc(ff)
        
        'bounds 2 (24 bytes)
        Echo " bounds 2 start at " & loc(ff)
        Get #ff, , .min2
        Get #ff, , .max2
        Echo " bounds 2 end at " & loc(ff)
        
        Echo "header end at " & loc(ff)
        Echo ""
        
        'meshnum (4 bytes)
        'Get #ff, , .meshnum
        'Echo "meshnum: " & .meshnum
        .meshnum = 4
        
        'meshes
        ReDim .mesh(0 To .meshnum - 1)
        For i = 0 To .meshnum - 1
            Echo "mesh " & i & " start at " & loc(ff)
            
            'matnum
            Get #ff, , .mesh(i).matnum
            Echo " matnum: " & .mesh(i).matnum
            
            'material groups
            If .mesh(i).matnum > 0 Then
                ReDim .mesh(i).mat(0 To .mesh(i).matnum - 1)
                For j = 0 To .mesh(i).matnum - 1
                    Echo " material " & i & " start at " & loc(ff)
                    
                    Get #ff, , .mesh(i).mat(j).start
                    Get #ff, , .mesh(i).mat(j).count
                    Echo "  start: " & .mesh(i).mat(j).start
                    Echo "  count: " & .mesh(i).mat(j).count
                    
                    .mesh(i).mat(j).texname = ReadTreeMeshString(ff)
                    Echo "  texname: " & .mesh(i).mat(j).texname
                    
                    Echo " material " & i & " end at " & loc(ff)
                Next j
            End If
            
            Echo "mesh " & i & " end at " & loc(ff)
            Echo ""
        Next i
        
        Echo ">>> collision data @ " & loc(ff)
        Echo ""
        
        'collision flag (4 bytes)
        Get #ff, , .colflag
        Echo "colflag: " & .colflag
        Echo ""
        
        If .colflag <> 0 Then
            Echo ">>> collision block start at " & loc(ff)
            
            'colu1 (4 bytes)
            Get #ff, , .colu1
            Echo " colu1: " & .colu1 'always 5? maybe chunk format version number
            
            '--- collider vertices ---------------------------------------------
            
            Echo " colvert block start at " & loc(ff)
            
            'colvertnum (4 bytes)
            Get #ff, , .colvertnum
            Echo "  colvertnum: " & .colvertnum
            
            'colvert (16 bytes * colvertnum)
            ReDim .colvert(0 To .colvertnum - 1)
            Get #ff, , .colvert()
            
            Echo " colvert block end at " & loc(ff)
            Echo ""
            
            '--- collider faces ------------------------------------------------
            
            Echo " colface block start at " & loc(ff)
            
            'colfacenum (4 bytes)
            Get #ff, , .colfacenum
            Echo "  colfacenum: " & .colfacenum
            
            'colface data (8 bytes * colfacenum)
            ReDim .colface(0 To .colfacenum - 1)
            Get #ff, , .colface()
            
            Echo " colface block end at " & loc(ff)
            Echo ""
            
            '--- unknown -----------------------------------------
            
            'h_u1 (4 bytes)
            Get #ff, , .h_u1 'always same as colfacenum?
            Echo " h_u1: " & .h_u1
            
            'h_u2 (4 bytes)
            Get #ff, , .h_u2 'always zero??
            Echo " h_u2: " & .h_u2
            
            Echo ""
            
            '--- h block ------------------------------------------
            
            Echo " h block start at " & loc(ff)
            
            'hnum (4 bytes)
            Get #ff, , .hnum 'always same as colfacenum?
            Echo "  hnum: " & .hnum
            
            'hdata (32 bytes * hnum)
            ReDim .hdata(0 To .hnum - 1)
            Get #ff, , .hdata()
            
            Echo " h block end at " & loc(ff)
            Echo ""
            
            '--- BSP tree -------------------------------------
            
            Dim dummy As tmaabb
            ReadTreeMeshNode ff, dummy
            
            '-----------------------------------------------
            
            Echo ""
            Echo ">>> collision block end at " & loc(ff)
            Echo ""
        End If
        
        '--- geometry data --------------------------------------------
        
        'vertexnum (4 bytes)
        Get #ff, , .vertnum
        Echo "vertnum: " & .vertnum
        
        'vertex data (44 bytes * vertnum)
        Echo "vert data start " & loc(ff)
        ReDim .vert(0 To .vertnum - 1)
        Get #ff, , .vert()
        Echo "vert data end " & loc(ff)
        
        'indexnum (4 bytes)
        Get #ff, , .indexnum
        Echo "indexnum: " & .indexnum
        
        'index data (2 bytes * indexnum)
        Echo "index data start " & loc(ff)
        ReDim .Index(0 To .indexnum - 1)
        Get #ff, , .Index()
        Echo "index data end " & loc(ff)
        
        '--- end of file ----------------------------------------------
        
        Echo "done reading " & loc(ff)
        Echo "file size is " & LOF(ff)
        Echo ""
        
        .loadok = True
        .drawok = True
        
        'generate colface normals
        If .colfacenum > 0 Then
            ReDim .colfacenorm(0 To .colfacenum - 1)
            For i = 0 To .colfacenum - 1
                Dim p1 As float3
                Dim p2 As float3
                Dim p3 As float3
                p1.x = .colvert(.colface(i).v1).x
                p1.y = .colvert(.colface(i).v1).y
                p1.z = .colvert(.colface(i).v1).z
                
                p2.x = .colvert(.colface(i).v2).x
                p2.y = .colvert(.colface(i).v2).y
                p2.z = .colvert(.colface(i).v2).z
                
                p3.x = .colvert(.colface(i).v3).x
                p3.y = .colvert(.colface(i).v3).y
                p3.z = .colvert(.colface(i).v3).z
                .colfacenorm(i) = GenNormal(p1, p2, p3)
            Next i
        End If
        
    End With
    
    'close file
    Close #ff
    
    LoadTreeMesh = True
    Exit Function
errorhandler:
    MsgBox "LoadTreeMesh" & vbLf & err.Description, vbCritical
    Echo ">>> error at " & loc(ff)
    Echo ">>> file size " & LOF(ff)
End Function


'reads collider AABB block
Private Function ReadTreeMeshNode(ByRef ff As Integer, ByRef node As tmaabb) As Boolean
    With node
        Echo ">>> aabb node @ " & loc(ff)
        
        'aabb (24 bytes)
        Get #ff, , .min
        Get #ff, , .max
        
        'facenum (4 bytes)
        Get #ff, , .facenum
        Echo ">>>  facenum: " & .facenum
        
        'face data (4 bytes * facenum)
        If .facenum > 0 Then
            ReDim .face(0 To .facenum - 1)
            Get #ff, , .face()
        End If
        
        Dim flag As Byte
        Dim dummy As tmaabb
        
        'A
        Get #ff, , flag
        If flag = 1 Then
            ReadTreeMeshNode ff, dummy
        End If
        
        'B
        Get #ff, , flag
        If flag = 1 Then
            ReadTreeMeshNode ff, dummy
        End If
        
    End With
    ReadTreeMeshNode = True
End Function


'reads string
Private Function ReadTreeMeshString(ByRef ff As Integer) As String
Dim num As Long
Dim chars() As Byte
    Get #ff, , num
    
    If num = 0 Then Exit Function
    
    ReDim chars(0 To num - 1)
    Get #ff, , chars()
    
    ReadTreeMeshString = SafeString(chars, num)
End Function


'draws TreeMesh
Public Sub DrawTreeMesh()
    On Error GoTo errorhandler
Dim i As Long
Dim j As Long
    
    With treemesh
        If Not .drawok Then Exit Sub
        
        'vertex pointer
        glVertexPointer 3, GL_FLOAT, 44, .vert(0).vert.x
        glNormalPointer GL_FLOAT, 44, .vert(0).norm.x
        glTexCoordPointer 2, GL_FLOAT, 44, .vert(0).tex.x
        
        If view_poly Then
            
            'draw solid
            For i = 0 To .meshnum - 1
                glEnableClientState GL_VERTEX_ARRAY
                glEnableClientState GL_NORMAL_ARRAY
                glEnableClientState GL_TEXTURE_COORD_ARRAY
                For j = 0 To .mesh(i).matnum - 1
                    
                    If view_textures Then
                        Dim texmapid As Long
                        texmapid = .mesh(i).mat(j).texmapid
                        If texmapid > 0 Then
                            glBindTexture GL_TEXTURE_2D, texmap(texmapid).tex
                            glEnable GL_TEXTURE_2D
                            
                            glColor3f 1, 1, 1
                            
                            Select Case i
                            Case 0 'leaf
                                'glBlendFunc GL_SRC_ALPHA, GL_ONE_MINUS_SRC_ALPHA
                                'glEnable GL_BLEND
                                'glDepthMask False
                                glDisable GL_CULL_FACE
                                
                                glEnable GL_ALPHA_TEST
                                glAlphaFunc GL_GREATER, 0.5
                                
                            Case 1 'trunk
                                '
                                
                            Case 2 'sprite
                                glBlendFunc GL_SRC_ALPHA, GL_ONE_MINUS_SRC_ALPHA
                                glEnable GL_BLEND
                                glDepthMask False
                                
                            Case 3 'unknown (maybe alpha-test?
                                'todo: research
                                
                            End Select
                            
                        Else
                            glColor3f 1, 0.25, 0.25
                        End If
                    Else
                        glColor3f 0.75, 0.75, 0.75
                    End If
                    If view_lighting Then
                        glEnable GL_LIGHTING
                    End If
                    If view_edges Or view_verts Then
                        glPolygonOffset 1, 1
                        glEnable GL_POLYGON_OFFSET_FILL
                    End If
                    
                    glDrawElements GL_TRIANGLES, .mesh(i).mat(j).count * 3, GL_UNSIGNED_SHORT, .Index(.mesh(i).mat(j).start)
                    
                    If view_edges Or view_verts Then
                        glDisable GL_POLYGON_OFFSET_FILL
                    End If
                    If view_lighting Then
                        glDisable GL_LIGHTING
                    End If
                    If view_textures Then
                        glDisable GL_TEXTURE_2D
                    End If
                    
                    glEnable GL_CULL_FACE
                    glDisable GL_BLEND
                    glDepthMask True
                    glDisable GL_ALPHA_TEST
                Next j
                glDisableClientState GL_TEXTURE_COORD_ARRAY
                glDisableClientState GL_NORMAL_ARRAY
                glDisableClientState GL_VERTEX_ARRAY
                
                If view_edges And Not view_wire Then
                    glEnableClientState GL_VERTEX_ARRAY
                    glColor4f 1, 1, 1, 0.1
                    StartAALine 1.3
                    glPolygonMode GL_FRONT_AND_BACK, GL_LINE
                    For j = 0 To .mesh(i).matnum - 1
                        glDrawElements GL_TRIANGLES, .mesh(i).mat(j).count * 3, GL_UNSIGNED_SHORT, .Index(.mesh(i).mat(j).start)
                    Next j
                    glPolygonMode GL_FRONT_AND_BACK, GL_FILL
                    EndAALine
                    glDisableClientState GL_VERTEX_ARRAY
                End If
            Next i
            
            'NOTE: bf42 treemesh files are bugged, some groups have 8x the number of indices!
            ''drawoutlines
            'If view_edges And Not view_wire Then
            '    glEnableClientState GL_VERTEX_ARRAY
            '    glColor4f 1, 1, 1, 0.1
            '    StartAALine 1.3
            '    glPolygonMode GL_FRONT_AND_BACK, GL_LINE
            '    glDrawElements GL_TRIANGLES, .indexnum, GL_UNSIGNED_SHORT, .index(0)
            '    glPolygonMode GL_FRONT_AND_BACK, GL_FILL
            '    EndAALine
            '    glDisableClientState GL_VERTEX_ARRAY
            'End If
        End If
        
        'draw vertices
        glEnableClientState GL_VERTEX_ARRAY
        If view_verts Then
            glColor3f 1, 1, 1
            StartAAPoint 4
            glDrawArrays GL_POINTS, 0, .vertnum
            EndAALine
        End If
        glDisableClientState GL_VERTEX_ARRAY
        
        
        'draw collider
        If view_bonesys Then
            
            'collider faces
            glColor3f 0.9, 0.8, 0.4
            'glPolygonOffset 1, 1
            'glEnable GL_POLYGON_OFFSET_FILL
            glEnable GL_LIGHTING
            glBegin GL_TRIANGLES
                For i = 0 To .colfacenum - 1
                    glNormal3fv .colfacenorm(i).x
                    glVertex3fv .colvert(.colface(i).v3).x
                    glVertex3fv .colvert(.colface(i).v2).x
                    glVertex3fv .colvert(.colface(i).v1).x
                Next i
            glEnd
            glDisable GL_LIGHTING
            'glDisable GL_POLYGON_OFFSET_FILL
            
            ''collider vertices
            'glColor3f 1, 0.9, 0.5
            'StartAAPoint 4
            'glBegin GL_POINTS
            '    For i = 0 To .colvertnum - 1
            '        glVertex3fv .colvert(i).x
            '    Next i
            'glEnd
            'EndAALine
            
        End If
        
        
        'draw normals
        If view_normals Then
            glColor3f 0, 1, 1
            StartAALine 1.3
            glBegin GL_LINES
            For i = 0 To .vertnum - 1
                With .vert(i)
                    Dim v2 As float3
                    v2.x = .vert.x + .norm.x * 0.05
                    v2.y = .vert.y + .norm.y * 0.05
                    v2.z = .vert.z + .norm.z * 0.05
                    glVertex3fv .vert.x
                    glVertex3fv v2.x
                End With
            Next i
            glEnd
            EndAALine
        End If
        
        'mesh bounds
        If view_bounds Then
            
            StartAALine 1.3
            
            'mesh aabb
            glColor3f 1, 1, 0
            DrawBox .min, .max
            
            glLineStipple 1, &HF0F
            glEnable GL_LINE_STIPPLE
            
            glColor3f 1, 0.5, 0
            DrawBox .min2, .max2
            
            'glColor3f 1, 0, 1
            'DrawBox .colmin, .colmax
            
            ''aabb tree
            'glColor3f 0, 1, 0
            'For i = 0 To .h_u2 - 1
            '    DrawBox .aabbnode(i).min, .aabbnode(i).max
            'Next i
            
            glDisable GL_LINE_STIPPLE
            
            EndAALine
            
        End If
        
    End With
    
    Exit Sub
errorhandler:
    MsgBox "DrawTreeMesh" & vbLf & err.Description, vbCritical
    treemesh.drawok = False
End Sub


'unloads StandardMesh
Public Sub UnloadTreeMesh()
    With treemesh
        .loadok = False
        .drawok = False
        .filename = ""
        
        .meshnum = 0
        Erase .mesh()
        
        .colvertnum = 0
        .colfacenum = 0
        Erase .colvert()
        Erase .colface()
        
        .vertnum = 0
        .indexnum = 0
        Erase .vert()
        Erase .Index
    End With
End Sub


'fills treeview hierarchy
Public Sub FillTreeTreeMesh(ByRef tree As MSComctlLib.TreeView)
    Dim i As Long
    Dim j As Long
    With treemesh
        If Not .loadok Then Exit Sub
        
        Dim n As MSComctlLib.node
        
        'file root
        Dim rootname As String
        rootname = "treemesh_root"
        Set n = tree.Nodes.Add(, , rootname, GetFileName(.filename), "file")
        n.Expanded = True
        n.tag = 0
        
        Set n = tree.Nodes.Add(rootname, tvwChild, "u1", "u1: " & .u1, "prop")
        Set n = tree.Nodes.Add(rootname, tvwChild, "u2", "u2: " & .u2, "prop")
        Set n = tree.Nodes.Add(rootname, tvwChild, "u3", "u3: " & .u3, "prop")
        
        Set n = tree.Nodes.Add(rootname, tvwChild, "vertnum", "Vertices: " & .vertnum, "prop")
        Set n = tree.Nodes.Add(rootname, tvwChild, "indexnum", "Indices: " & .indexnum, "prop")
        
        If .meshnum > 0 Then
            
            Dim tag As Long
            
            Dim meshname As String
            meshname = "mesh"
            Set n = tree.Nodes.Add(rootname, tvwChild, meshname, "Mesh", "geom")
            n.Expanded = True
            n.tag = MakeTag(0, 0, 0)
            
            'lods
            For i = 0 To .meshnum - 1
                With .mesh(i)
                
                    Dim lodname As String
                    lodname = "lod " & i
                    tag = MakeTag(0, i, 0)
                    
                    Set n = tree.Nodes.Add(meshname, tvwChild, lodname, "Group " & i, "lod")
                    n.tag = tag
                    n.Expanded = True
                    
                    'Set n = tree.Nodes.Add(lodname, tvwChild, lodname & "|trinum", "Polygons: " & .polycount, "trinum")
                    'n.tag = tag
                    
                    'materials
                    For j = 0 To .matnum - 1
                        With .mat(j)
                            
                            Dim matname As String
                            matname = lodname & "|mat " & j
                            Set n = tree.Nodes.Add(lodname, tvwChild, matname, "Mat " & j, "mat")
                            n.tag = tag
                            n.Expanded = True
                            
                            Set n = tree.Nodes.Add(matname, tvwChild, matname & "|start", "Start: " & .start, "prop")
                            n.tag = tag
                            
                            Set n = tree.Nodes.Add(matname, tvwChild, matname & "|trinum", "Triangles: " & .count, "trinum")
                            n.tag = tag
                            
                            Dim texicon As String
                            If .texmapid = 0 Then
                                texicon = "texmissing"
                            Else
                                texicon = "tex"
                            End If
                            
                            Set n = tree.Nodes.Add(matname, tvwChild, matname & "|tex", .texname, texicon)
                            n.tag = tag
                            
                        End With
                    Next j
                End With
            Next i
        End If
        
        If .u1 > 0 Then
            'meshname = "colmesh"
            'Set n = tree.Nodes.Add(rootname, tvwChild, meshname, "Collision Meshes", "geom")
            'n.Expanded = True
            'n.tag = MakeTag(1, 0)
            
            'cols
            'For i = 0 To .colnum - 1
            '    With .col(i)
            '
            '        lodname = "col " & i
            '        tag = MakeTag(1, i)
            '
            '        Set n = tree.Nodes.Add(meshname, tvwChild, lodname, "Col " & i, "lod")
            '        n.tag = tag
            '
            '        Set n = tree.Nodes.Add(lodname, tvwChild, lodname & "|trinum", "Polygons: " & .facenum, "trinum")
            '        n.tag = tag
            '
            '        Set n = tree.Nodes.Add(lodname, tvwChild, lodname & "|vertnum", "Vertices: " & .vertnum, "prop")
            '        n.tag = tag
            '
            '    End With
            'Next i
        End If
        
    End With
End Sub

