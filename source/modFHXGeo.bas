Attribute VB_Name = "FHX_Geo"
Option Explicit

Private Type byte4
    r As Byte
    g As Byte
    b As Byte
    a As Byte
End Type

Private Type skinbone
    nodeid As Long
    pos As float3
    rot As quat
    
    'internal
    mat As matrix4
End Type

Private Type matgroup       '20 bytes
    matid As Long
    indexcount As Long
    indexstart As Long
    indexrangemin As Long
    indexrangemax As Long
End Type

Private Type geomlod        '16 bytes + ...
    
    'vertex info
    vertnum As Long
    tccnum As Long
    tangentChannels As Long
    hasnormals As Byte
    hascolors As Byte
    bonesPerVert As Long
    
    'vertex bounds '24 bytes
    min As float3
    max As float3
    
    'bone system
    bonenum As Long
    bone() As skinbone
    
    'vertex data
    vert() As float3
    texc() As float2
    norm() As byte4
    tan1() As byte4
    tan2() As byte4
    color() As byte4
    weight() As Byte
    boneId() As Byte
    
    'indices
    indexnum As Long
    Index() As Integer
    
    'material groups
    matgroupnum As Long
    matgroup() As matgroup
    
    'internal
    inorm() As float3
    itan1() As float3 'tangent channel one S
    itan2() As float3 'tangent channel one T
    vertflag() As Boolean      'array of vertex draw flags (to prevent overdraw of vertices, normals etc)
End Type

Private Type fhxgeo_type
    head As fileheader
    lodnum As Long
    lod() As geomlod
    
    'internal
    filename As String
    loadok As Boolean
    drawok As Boolean
End Type

Public fhxgeo As fhxgeo_type


'reads FHX geometry file
Public Function LoadFhxGeo(ByVal filename As String)
    On Error GoTo errhandler
    
Dim i As Long
Dim j As Long
    
    'open file
    Dim ff As Integer
    ff = FreeFile
    Open filename For Binary Access Read Lock Write As #ff
    
    With fhxgeo
        .loadok = True
        .filename = filename
        
        'read header
        Get #ff, , .head
        Echo "file format version: " & .head.version
        Echo "size: " & .head.size & "/" & LOF(ff)
        Echo "offset: " & .head.offset
        Echo ""
        
        If .head.version < 6 Then
            MsgBox "File version no longer supported.", vbExclamation
            Close #ff
            Exit Function
        End If
        
        'read number of LODs
        Get #ff, , .lodnum
        Echo "lodnum: " & .lodnum
        
        'read LODs
        ReDim .lod(0 To .lodnum - 1)
        For i = 0 To .lodnum - 1
            Echo "lod " & i & " start at " & loc(ff)
            fhxReadLod ff, .lod(i), .head.version
            Echo "lod " & i & " end at " & loc(ff)
            Echo ""
        Next i
        
        'done
        Echo "done reading " & loc(ff)
        Echo "file size is " & LOF(ff)
        Echo ""
        
        If loc(ff) <> LOF(ff) Then
            MsgBox "File size does not match, please check for data corruption.", vbExclamation
        End If
        
        .loadok = True
        .drawok = True
    End With
    
    'close file
    Close ff
    
    'success
    LoadFhxGeo = True
    Exit Function
    
    'error handler
errhandler:
    MsgBox "fhxLoadGeo" & vbLf & err.Description, vbCritical
End Function


'reads LOD
Private Sub fhxReadLod(ByRef ff As Integer, ByRef lod As geomlod, ByVal version As Long)
Dim i As Long
    With lod
        
        Dim tmp As Byte
        
        'read vertex attribute info
        Get #ff, , .vertnum
        Get #ff, , .tccnum
        If version >= 7 Then
            Get #ff, , .tangentChannels
        End If
        Get #ff, , .hasnormals
        If version < 7 Then
            Get #ff, , tmp
            If tmp > 0 Then
                .tangentChannels = 2
            Else
                .tangentChannels = 0
            End If
        End If
        Get #ff, , .hascolors
        If version < 7 Then
            Get #ff, , tmp
            .bonesPerVert = tmp
        Else
            Get #ff, , .bonesPerVert
        End If
        Echo "header end @ " & loc(ff)
        
        Echo " vertnum: " & .vertnum
        Echo " tccnum: " & .tccnum
        Echo " tangentChannels: " & .tangentChannels
        Echo " hasnormals: " & .hasnormals
        Echo " hascolors: " & .hascolors
        Echo " bonesPerVert: " & .bonesPerVert
        Echo ""
        
        'read bounds
        Echo " bounds at " & loc(ff)
        Get #ff, , .min
        Get #ff, , .max
        
        'read bone system
        Get #ff, , .bonenum
        Echo " bonenum: " & .bonenum
        If .bonenum > 0 Then
            ReDim .bone(0 To .bonenum - 1)
            Echo " bone data start at " & loc(ff)
            For i = 0 To .bonenum - 1
                With .bone(i)
                    Get #ff, , .nodeid
                    Get #ff, , .pos
                    Get #ff, , .rot
                    
                    Echo " bone " & i & ": id:" & .nodeid ' & " parent: " & .parent
                
                    mat4identity .mat
                    mat4setpos .mat, .pos
                    mat4setrot .mat, .rot
                End With
                'If .bone(i).parent > -1 Then
                '    .bone(i).mat = mat4mult(.bone(i).mat, .bone(.bone(i).parent).mat)
                'End If
            Next i
            Echo " bone data end at " & loc(ff)
        End If
        
        'read vertices
        If .vertnum > 0 Then
            Echo " vertices at " & loc(ff)
            ReDim .vert(0 To .vertnum - 1)
            Get #ff, , .vert()
        End If
        
        'read texcoords
        If .tccnum > 0 Then
            Echo " texcoords at " & loc(ff)
            ReDim .texc(0 To (.vertnum * .tccnum) - 1)
            Get #ff, , .texc()
        End If
        
        'read normals
        If .hasnormals > 0 Then
            Echo " normals at " & loc(ff)
            ReDim .norm(0 To .vertnum - 1)
            Get #ff, , .norm()
        End If
        
        'read tangents
        If .tangentChannels > 0 Then
            Echo " tangents at " & loc(ff)
            
            If .tangentChannels >= 1 Then
                ReDim .tan1(0 To .vertnum - 1)
                Get #ff, , .tan1()
            End If
            
            If .tangentChannels >= 2 Then
                ReDim .tan2(0 To .vertnum - 1)
                Get #ff, , .tan2()
            End If
        End If
        
        'read colors
        If .hascolors > 0 Then
            Echo " colors at " & loc(ff)
            ReDim .color(0 To .vertnum - 1)
            Get #ff, , .color()
        End If
        
        'read weights
        If .bonesPerVert > 1 Then
            Echo " weights at " & loc(ff)
            ReDim .weight(0 To (.bonesPerVert * .vertnum) - 1)
            Get #ff, , .weight()
        End If
        
        'read boneids
        If .bonesPerVert > 0 Then
            Echo " boneids at " & loc(ff)
            ReDim .boneId(0 To (.bonesPerVert * .vertnum) - 1)
            Get #ff, , .boneId()
            
            'For i = 0 To .vertnum - 1
            '    Echo "boneid " & .boneid((i * .bonesPerVert) + 0) & " " & _
            '                     .boneid((i * .bonesPerVert) + 1) & " " & _
            '                     .boneid((i * .bonesPerVert) + 2) & " " & _
            '                     .boneid((i * .bonesPerVert) + 3)
            'Next i
        End If
        
        'read indices
        Get #ff, , .indexnum
        Echo " indexnum: " & .indexnum
        If .indexnum > 0 Then
            ReDim .Index(0 To .indexnum - 1)
            Get #ff, , .Index()
        End If
        
        'read material groups
        Get #ff, , .matgroupnum
        Echo " matgroupnum: " & .matgroupnum
        If .matgroupnum > 0 Then
            ReDim .matgroup(0 To .matgroupnum - 1)
            For i = 0 To .matgroupnum - 1
                With .matgroup(i)
                    Echo ""
                    Echo " matgroup start at " & i & " at " & loc(ff)
                    
                    Get #ff, , .matid
                    Get #ff, , .indexcount
                    Get #ff, , .indexstart
                    Get #ff, , .indexrangemin
                    Get #ff, , .indexrangemax
                    
                    Echo "  matid: " & .matid
                    Echo "  indexcount: " & .indexcount
                    Echo "  indexstart: " & .indexstart
                    Echo "  indexrangemin: " & .indexrangemin
                    Echo "  indexrangemax: " & .indexrangemax
                    
                    Echo " matgroup end at " & i & " at " & loc(ff)
                End With
            Next i
        End If
        
        '--- internal stuff ---------------------------------------------------
        
        'convert byte4 to float3 for easier use
        If .hasnormals Then
            ReDim .inorm(0 To .vertnum - 1)
            For i = 0 To .vertnum - 1
                .inorm(i).x = ((CSng(.norm(i).r) / 255) - 0.5) * 2
                .inorm(i).y = ((CSng(.norm(i).g) / 255) - 0.5) * 2
                .inorm(i).z = ((CSng(.norm(i).b) / 255) - 0.5) * 2
            Next i
        End If
        If .tangentChannels > 0 Then
            ReDim .itan1(0 To .vertnum - 1)
            ReDim .itan2(0 To .vertnum - 1)
            For i = 0 To .vertnum - 1
                .itan1(i).x = ((CSng(.tan1(i).r) / 255) - 0.5) * 2
                .itan1(i).y = ((CSng(.tan1(i).g) / 255) - 0.5) * 2
                .itan1(i).z = ((CSng(.tan1(i).b) / 255) - 0.5) * 2
                
                Dim flip As Single
                flip = ((CSng(.tan1(i).a) / 255) - 0.5) * 2
                .itan2(i) = ScaleFloat3(CrossProduct(.inorm(i), .itan1(i)), flip)
                
                '.itan2(i).x = ((CSng(.tan2(i).r) / 255) - 0.5) * 2
                '.itan2(i).y = ((CSng(.tan2(i).g) / 255) - 0.5) * 2
                '.itan2(i).z = ((CSng(.tan2(i).b) / 255) - 0.5) * 2
            Next i
        End If
        
        'allocate vertflags
        ReDim .vertflag(0 To .vertnum - 1)
        
    End With
End Sub


'draws geometry
Public Sub DrawFhxGeo()
    If Not fhxgeo.loadok Then Exit Sub
    If Not fhxgeo.drawok Then Exit Sub
    On Error GoTo errorhandler
    
    Dim i As Long
    Dim j As Long
    
    Dim v As float3
    Dim n As float3
    Dim t1 As float3
    Dim t2 As float3
    
    With fhxgeo
        With .lod(sellod)
            
            'compute normal/tangent vector drawing length
            Dim vecscale As Single
            vecscale = .max.x - .min.x
            vecscale = vecscale + (.max.y - .min.x)
            vecscale = vecscale + (.max.z - .min.z)
            vecscale = vecscale / 3
            vecscale = vecscale * 0.01
            vecscale = max(0.005, vecscale)
            
            'draw faces
            If view_poly Then
                
                glDisable GL_LIGHTING
                glDisable GL_TEXTURE_2D
                glDisable GL_BLEND
                glDisable GL_ALPHA_TEST
                
                'draw solid
                If .hasnormals And view_lighting Then
                    glEnable GL_LIGHTING
                End If
                glPolygonOffset 1, 1
                glEnable GL_POLYGON_OFFSET_FILL
                
                glColor3f 0.75, 0.75, 0.75
                
                'drawcall
                glEnableClientState GL_VERTEX_ARRAY
                glVertexPointer 3, GL_FLOAT, 0, .vert(0).x
                
                If .hasnormals Then
                    glEnableClientState GL_NORMAL_ARRAY
                    glNormalPointer GL_FLOAT, 0, .inorm(0).x
                    'glEnableClientState GL_COLOR_ARRAY
                    'glColorPointer 4, GL_BYTE, 0, .norm(0).r
                End If
                If .hascolors Then
                    glEnableClientState GL_COLOR_ARRAY
                    glColorPointer 4, GL_BYTE, 0, .color(0).r
                End If
                'glEnableClientState GL_COLOR_ARRAY
                'glColorPointer 4, GL_BYTE, 0, .tan2(0).r
                
                glDrawElements GL_TRIANGLES, .indexnum, GL_UNSIGNED_SHORT, .Index(0)
                
                glDisableClientState GL_COLOR_ARRAY
                glDisableClientState GL_NORMAL_ARRAY
                glDisableClientState GL_VERTEX_ARRAY
                
                glDisable GL_POLYGON_OFFSET_FILL
                glDisable GL_LIGHTING
                
                'draw edges
                If view_edges And Not view_wire Then
                    glColor4f 1, 1, 1, 0.1
                    StartAALine 1.3
                    glVertexPointer 3, GL_FLOAT, 0, .vert(0).x
                    
                    glPolygonMode GL_FRONT_AND_BACK, GL_LINE
                    glEnableClientState GL_VERTEX_ARRAY
                    
                    glDrawElements GL_TRIANGLES, .indexnum, GL_UNSIGNED_SHORT, .Index(0)
                    
                    glDisableClientState GL_VERTEX_ARRAY
                    glPolygonMode GL_FRONT_AND_BACK, GL_FILL
                    EndAALine
                End If
                
            End If
            
            'draw UVs
            If 1 = 2 Then
            If .tccnum > 0 Then
                glColor3f 0.5, 0.75, 1
                StartAALine 1.3
                
                glVertexPointer 2, GL_FLOAT, 0, .texc(0).x
                    
                glPolygonMode GL_FRONT_AND_BACK, GL_LINE
                glEnableClientState GL_VERTEX_ARRAY
                
                glDrawElements GL_TRIANGLES, .indexnum, GL_UNSIGNED_SHORT, .Index(0)
                
                glDisableClientState GL_VERTEX_ARRAY
                glPolygonMode GL_FRONT_AND_BACK, GL_FILL
                
                EndAALine
            End If
            End If
            
            'clear vert flags
            For i = 0 To .vertnum - 1
                .vertflag(i) = False
            Next i
            
            'generate vertex flags
            For i = 0 To .indexnum - 1
                .vertflag(.Index(i)) = True
            Next i
            
            'draw vertices
            If view_verts Then
                glColor4f 1, 1, 1, 1
                StartAAPoint 4
                glBegin GL_POINTS
                For i = 0 To .vertnum - 1
                    If .vertflag(i) Then
                        glVertex3fv .vert(i).x
                    End If
                Next i
                glEnd
                EndAAPoint
            End If
            
            'draw tangents
            If .tangentChannels > 0 And view_tangents Then
                StartAALine 1.3
                glBegin GL_LINES
                    For i = 0 To .vertnum - 1
                        If .vertflag(i) Then
                            v = .vert(i)
                            t1 = AddFloat3(v, ScaleFloat3(.itan1(i), vecscale))
                            t2 = AddFloat3(v, ScaleFloat3(.itan2(i), vecscale))
                            
                            'draw tangent
                            glColor4f 1, 0.5, 0.5, 0.5 'red
                            glVertex3fv v.x
                            glVertex3fv t1.x
                            
                            'draw bitangent
                            glColor4f 0.5, 1, 0.5, 0.5 'green
                            glVertex3fv v.x
                            glVertex3fv t2.x
                        End If
                    Next i
                glEnd
                EndAALine
            End If
            
            'draw normals
            If .hasnormals And view_normals Then
                StartAALine 1.3
                glColor4f 0, 1, 1, 0.5 'cyan
                For i = 0 To .vertnum - 1
                    If .vertflag(i) Then
                        glBegin GL_LINES
                            v = .vert(i)
                            n = AddFloat3(v, ScaleFloat3(.inorm(i), vecscale))
                            glVertex3fv v.x
                            glVertex3fv n.x
                        glEnd
                    End If
                Next i
                EndAALine
            End If
            
            'bone system
            If view_bonesys Then
                
                'draw vert links
                If .bonesPerVert > 0 Then
                    StartAALine 1.3
                    
                    Dim b1 As Long
                    Dim b2 As Long
                    Dim b3 As Long
                    Dim b4 As Long
                    Dim w1 As Byte
                    Dim w2 As Byte
                    Dim w3 As Byte
                    Dim w4 As Byte
                    
                    glBegin GL_LINES
                    For i = 0 To .vertnum - 1
                        
                        If .bonesPerVert >= 1 Then b1 = .boneId(i * .bonesPerVert + 0)
                        If .bonesPerVert >= 2 Then b2 = .boneId(i * .bonesPerVert + 1)
                        If .bonesPerVert >= 3 Then b3 = .boneId(i * .bonesPerVert + 2)
                        If .bonesPerVert = 4 Then b4 = .boneId(i * .bonesPerVert + 3)
                        
                        If .bonesPerVert = 1 Then w1 = 1
                        If .bonesPerVert >= 2 Then w1 = .weight(i * .bonesPerVert + 0)
                        If .bonesPerVert >= 2 Then w2 = .weight(i * .bonesPerVert + 1)
                        If .bonesPerVert >= 3 Then w3 = .weight(i * .bonesPerVert + 2)
                        If .bonesPerVert = 4 Then w4 = .weight(i * .bonesPerVert + 3)
                        
                        If w1 > 0 Then
                            glColor4f 1, 0.5, 0.5, 0.1
                            glVertex3fv .vert(i).x
                            glVertex3fv .bone(b1).mat.m(12)
                        End If
                        
                        If w2 > 0 Then
                            glColor4f 0.5, 0.5, 1, 0.1
                            glVertex3fv .vert(i).x
                            glVertex3fv .bone(b2).mat.m(12)
                        End If
                        
                        If w3 > 0 Then
                            glColor4f 0.5, 1, 0.5, 0.1
                            glVertex3fv .vert(i).x
                            glVertex3fv .bone(b3).mat.m(12)
                        End If
                        
                        If w4 > 0 Then
                            glColor4f 1, 1, 0.5, 0.1
                            glVertex3fv .vert(i).x
                            glVertex3fv .bone(b4).mat.m(12)
                        End If
                        
                    Next i
                    glEnd
                    
                    EndAALine
                End If
                
                'draw bones
                glColor3f 1, 1, 0
                StartAAPoint 5
                glDisable GL_DEPTH_TEST
                For i = 0 To .bonenum - 1
                    
                    'draw dot
                    glBegin GL_POINTS
                        glVertex3fv .bone(i).mat.m(12)
                    glEnd
                    
                Next i
                glEnable GL_DEPTH_TEST
                EndAAPoint
                
            End If
            
            'draw bounds
            If view_bounds Then
                StartAALine 1.3
                glColor3f 1, 1, 0
                DrawBox .min, .max
                EndAALine
            End If
            
            
        End With
                
    End With
    
    Exit Sub
errorhandler:
    MsgBox "fhxDrawGeo()" & err.Description, vbCritical
    fhxgeo.drawok = False
End Sub


'unloads geometry
Public Sub UnloadFhxGeo()
    With fhxgeo
        .loadok = False
        .drawok = False
        .filename = ""
        
        .lodnum = 0
        Erase .lod()
    End With
End Sub

Public Sub FillTreeFhxGeo(ByRef tree As MSComctlLib.TreeView)
    Dim i As Long
    Dim j As Long
    With fhxgeo
        If Not .loadok Then Exit Sub
        
        Dim n As MSComctlLib.node
        
        'file root
        Dim rootname As String
        rootname = "fhx_geo"
        Set n = tree.Nodes.Add(, tvwChild, rootname, GetFileName(.filename), "file")
        n.Expanded = True
        n.tag = 0
        
        'version
        Set n = tree.Nodes.Add(rootname, tvwChild, rootname & "|ver", "Version: " & .head.version, "prop")
        n.tag = 0
        
        'lods
        For i = 0 To fhxgeo.lodnum - 1
            With fhxgeo.lod(i)
                
                'add lod
                Dim lodname As String
                lodname = "lod " & i
                Set n = tree.Nodes.Add(rootname, tvwChild, lodname, "Lod " & i, "lod")
                n.Expanded = True
                n.tag = i
                
                'add lod properties
                
                Set n = tree.Nodes.Add(lodname, tvwChild, lodname & "|trinum", "Polygons: " & (.indexnum / 3), "trinum")
                n.tag = i
                
                Set n = tree.Nodes.Add(lodname, tvwChild, lodname & "|vertnum", "Vertices: " & .vertnum, "prop")
                n.tag = i
                
                Set n = tree.Nodes.Add(lodname, tvwChild, lodname & "|uvnum", "UV Channels: " & .tccnum, "prop")
                n.tag = i
                
                Set n = tree.Nodes.Add(lodname, tvwChild, lodname & "|hasnorm", "Normals: " & YesNo(.hasnormals), "prop")
                n.tag = i
                
                Set n = tree.Nodes.Add(lodname, tvwChild, lodname & "|hastan", "Tangents: " & .tangentChannels, "prop")
                n.tag = i
                
                Set n = tree.Nodes.Add(lodname, tvwChild, lodname & "|hasclr", "Colors: " & YesNo(.hascolors), "prop")
                n.tag = i
                
                Set n = tree.Nodes.Add(lodname, tvwChild, lodname & "|bpv", "Bones Per Vertex: " & .bonesPerVert, "prop")
                n.tag = i
                
                Set n = tree.Nodes.Add(lodname, tvwChild, lodname & "|indnum", "Indices: " & .indexnum, "prop")
                n.tag = i
                
                Dim bonesysname As String
                bonesysname = lodname & "|bonesys"
                Set n = tree.Nodes.Add(lodname, tvwChild, bonesysname, "Bones: " & .bonenum, "prop")
                n.tag = i
                
                For j = 0 To .bonenum - 1
                    Set n = tree.Nodes.Add(bonesysname, tvwChild, bonesysname & "|bone" & j, "NodeID " & .bone(j).nodeid, "prop")
                    n.tag = i
                Next j
                
                'materials
                For j = 0 To .matgroupnum - 1
                    With .matgroup(j)
                        
                        Dim matname As String
                        matname = lodname & "|mat " & j
                        Set n = tree.Nodes.Add(lodname, tvwChild, matname, "Mat " & j, "mat")
                        'n.Expanded = True
                        n.tag = i
                        
                        Set n = tree.Nodes.Add(matname, tvwChild, matname & "|matid", "Mat ID: " & .matid, "prop")
                        n.tag = i
                        
                        Set n = tree.Nodes.Add(matname, tvwChild, matname & "|trinum", "Polygons: " & (.indexcount / 3), "trinum")
                        n.tag = i
                        
                        Set n = tree.Nodes.Add(matname, tvwChild, matname & "|indcnt", "Indices: " & .indexcount, "prop")
                        n.tag = i
                        
                        Set n = tree.Nodes.Add(matname, tvwChild, matname & "|indsrt", "Index Start: " & .indexstart, "prop")
                        n.tag = i
                        
                        Set n = tree.Nodes.Add(matname, tvwChild, matname & "|indmin", "Min Vertex: " & .indexrangemin, "prop")
                        n.tag = i
                        
                        Set n = tree.Nodes.Add(matname, tvwChild, matname & "|indmax", "Max Vertex: " & .indexrangemax, "prop")
                        n.tag = i
                        
                    End With
                Next j
            End With
        Next i
        
    End With
End Sub

