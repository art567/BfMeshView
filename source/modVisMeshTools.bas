Attribute VB_Name = "BF2_MeshTools"
Option Explicit


'gets vertex weight
Public Sub GetSkinVertWeight(ByVal i As Long, ByRef vw As bf2skinweight)
    CopyMem VarPtr(vw), VarPtr(vmesh.vert(i * vmesh.xstride + 6)), 8
End Sub

'sets vertex weight
Public Sub SetSkinVertWeight(ByVal i As Long, ByRef vw As bf2skinweight)
    CopyMem VarPtr(vmesh.vert(i * vmesh.xstride + 6)), VarPtr(vw), 8
End Sub


'orients normal/tangent vectors to worldspace
Public Sub VisMeshTool_MakeWS()
Dim i As Long
Dim stride As Long
Dim normoff As Long
Dim tangoff As Long

    With vmesh
        
        stride = .vertstride / 4
        normoff = 3
        tangoff = ((.vertstride - 24) / 4) + 3
        
        For i = 0 To .vertnum - 1
            
            .vert((i * stride) + normoff + 0) = 1
            .vert((i * stride) + normoff + 1) = 0
            .vert((i * stride) + normoff + 2) = 0
            
            .vert((i * stride) + tangoff + 0) = 0
            .vert((i * stride) + tangoff + 1) = 1
            .vert((i * stride) + tangoff + 2) = 0
            
        Next i
        
    End With
End Sub


'computes vegetation normals
Public Sub VisMeshTool_VeggieNormals()
Dim i As Long
Dim j As Long
Dim g As Long
Dim m As Long
    With vmesh
        
        'compute vertex stride
        Dim stride As Long
        stride = .vertstride / 4
        
        'compute bounding box
        For g = 0 To .geomnum - 1
            With .geom(g)
                For j = 0 To .lodnum - 1
                    With .lod(j)
                        For m = 0 To .matnum - 1
                            With .mat(m)
                                If .technique = "Base" Then
                                    
                                    'todo <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
                                    
                                    'construct line segment
                                    's.x = .mmax.x - .mmin.x
                                    's.y = .mmax.y - .mmin.y
                                    's.z = .mmax.z - .mmin.z
                                    
                                    'Dim smin As Single
                                    'smin = min(min(s.x, szy), s.z)
                                    
                                    's.x = s.x - smin
                                    's.y = s.y - smin
                                    's.z = s.z - smin
                                    
                                    Dim p1 As float3
                                    Dim p2 As float3
                                    
                                    'p1.x = ((.mmax.x + .mmin.x) / 2) + (s.x / 2)
                                    'p1.y = ((.mmax.y + .mmin.y) / 2) + (s.y / 2)
                                    'p1.z = ((.mmax.z + .mmin.z) / 2) + (s.z / 2)
                                    
                                    'p2.x = ((.mmax.x + .mmin.x) / 2) + (s.x / 2)
                                    'p2.y = ((.mmax.y + .mmin.y) / 2) + (s.y / 2)
                                    'p2.z = ((.mmax.z + .mmin.z) / 2) + (s.z / 2)
                                    
                                    For i = .vstart To .vstart + .vnum - 1
                                        MakeVeggieNormal i * stride, p1, p2
                                    Next i
                                    
                                End If
                            End With
                        Next m
                    End With
                Next j
            End With
        Next g
    End With
End Sub

Private Sub MakeVeggieNormal(ByVal i As Long, ByRef p1 As float3, ByRef p2 As float3)
Dim v As float3
Dim n As float3
    With vmesh
        
        'get vertex
        v.x = .vert(i + 0)
        v.y = .vert(i + 1)
        v.z = .vert(i + 2)
        
        'compute normal
        'n = Normalize( ClosestPointOnLine(p1, p2, v) )
        n = Normalize(v)
        
        'replace old normal
        .vert(i + 3 + 0) = n.x
        .vert(i + 3 + 1) = n.y
        .vert(i + 3 + 2) = n.z
        
    End With
End Sub


' returns closest point on a line
Private Function ClosestPointOnLine(vA As float3, vB As float3, vPoint As float3) As float3

    ' Create the vector from end point vA to our point vPoint.
    Dim vVector1 As float3
    vVector1 = SubFloat3(vPoint, vA)
    
    ' Create a normalized direction vector from end point vA to end point vB
    Dim vVector2 As float3
    vVector2 = Normalize(SubFloat3(vB, vA))
    
    ' Use the distance formula to find the distance of the line segment (or magnitude)
    Dim d As Single
    d = Distance(vA, vB)
    
    ' Using the dot product, we project the vVector1 onto the vector vVector2.
    ' This essentially gives us the distance from our projected vector from vA.
    Dim t As Single
    t = DotProduct(vVector2, vVector1)
    
    ' If our projected distance from vA, "t", is less than or equal to 0, it must
    ' be closest to the end point vA.  We want to return this end point.
    If (t <= 0) Then
        ClosestPointOnLine = vA
        Exit Function
    End If
    
    ' If our projected distance from vA, "t", is greater than or equal to the magnitude
    ' or distance of the line segment, it must be closest to the end point vB.  So, return vB.
    If (t >= d) Then
        ClosestPointOnLine = vB
        Exit Function
    End If
    
    'Here we create a vector that is of length t and in the direction of vVector2
    Dim vVector3 As float3
    vVector3 = ScaleFloat3(vVector2, t)
    
    'To find the closest point on the line segment, we just add vVector3 to the original end point vA.
    ClosestPointOnLine = AddFloat3(vA, vVector3)
End Function


'verifies BF2 mesh
Public Sub BF2VerifyMesh()
    With vmesh
        If Not .loadok Then
            MsgBox "No BF2 mesh not loaded.", vbExclamation
            Exit Sub
        End If
        
        Dim i As Long
        Dim errstr As String
        
        'check vertex buffer for NaNs
        Dim vnum As Long
        Dim badvert As Long
        vnum = (.vertstride / .vertformat) * .vertnum
        For i = 0 To vnum - 1
            If IsNaN(.vert(i)) Then
                .vert(i) = 0
                badvert = badvert + 1
            End If
        Next i
        If badvert > 0 Then
            errstr = errstr & "* fixed " & badvert & " bad vertices"
        End If
        
        'check faces for thin triangles
        Dim g As Long
        For g = 0 To .geomnum - 1
            Dim L As Long
            For L = 0 To .geom(g).lodnum - 1
                
                'check bounds for NaNs
                Dim nan As Long
                nan = 0
                If IsNaN3f(.geom(g).lod(L).min) Then nan = nan + 1
                If IsNaN3f(.geom(g).lod(L).max) Then nan = nan + 1
                If nan > 0 Then
                    errstr = errstr & "* NaNs in LOD bounds"
                End If
                
                'verify materials
                Dim m As Long
                For m = 0 To .geom(g).lod(L).matnum - 1
                    BF2VerifyMat .geom(g).lod(L).mat(m)
                    
                    nan = 0
                    If IsNaN3f(.geom(g).lod(L).mat(m).mmin) Then nan = nan + 1
                    If IsNaN3f(.geom(g).lod(L).mat(m).mmin) Then nan = nan + 1
                    If nan > 0 Then
                        errstr = errstr & "* NaNs in material bounds"
                    End If
                    
                Next m
            Next L
        Next g
        
        'show stats
        If Len(errstr) > 0 Then
            MsgBox errstr, vbExclamation
        Else
            MsgBox "All ok!", vbInformation
        End If
        
    End With
End Sub


'counts number of degenerate triangles
Private Sub BF2VerifyMat(ByRef mat As bf2mat)
Dim i As Long
Dim v1 As float3
Dim v2 As float3
Dim v3 As float3
Dim a1 As Single
Dim a2 As Single
Dim a3 As Single
Dim err As Boolean
    
    With mat
        Dim badtri As Long
        badtri = 0
        
        'For i = 0 To .facenum - 1
            
            'err = False
            
            'v1 = .vert(.face(i).v1)
            'v2 = .vert(.face(i).v2)
            'v3 = .vert(.face(i).v3)
            
            'a1 = AngleBetweenVectors(SubFloat3(v1, v2), SubFloat3(v1, v3))
            'a2 = AngleBetweenVectors(SubFloat3(v2, v3), SubFloat3(v2, v1))
            'a3 = AngleBetweenVectors(SubFloat3(v3, v1), SubFloat3(v3, v2))
            
            'DEGENERATEFACEANGLE
            'Const badangle As Single = 0.1
            
            'If a1 < badangle Then err = True
            'If a2 < badangle Then err = True
            'If a3 < badangle Then err = True
            
            'If err Then
            '    badtri = badtri + 1
            'End If
        'Next i
        
    End With
End Sub


'make texture path BF2-friendly
' * replaces slashes
' * makes filename lowercase
Private Function FixTexPath(ByVal path As String)
    If InStr(1, LCase(path), "specularlut_pow36") Then
        FixTexPath = "Common\Textures\SpecularLUT_pow36.dds"
    Else
        FixTexPath = LCase(Replace(path, "\", "/"))
    End If
End Function


'fixes texture paths
Public Sub BF2MeshFixTexPaths()
    With vmesh
        If Not .loadok Then Exit Sub
        
        Dim g As Long
        For g = 0 To .geomnum - 1
            With .geom(g)
            
                Dim L As Long
                For L = 0 To .lodnum - 1
                     With .lod(L)
                        
                        Dim m As Long
                        For m = 0 To .matnum - 1
                            With .mat(m)
                                
                                Dim t As Long
                                For t = 0 To .mapnum - 1
                                    
                                    .map(t) = FixTexPath(.map(t))
                                    
                                Next t
                                
                            End With
                        Next m
                        
                     End With
                Next L
                
            End With
        Next g
        
    End With
    
    LoadMeshTextures
    frmMain.FillTreeView
End Sub


Public Sub PasteMaterial(ByRef dst As bf2mat, ByRef src As bf2mat)
    Dim i As Long
    With dst
        
        dst.alphamode = src.alphamode
        dst.fxfile = src.fxfile
        dst.technique = src.technique
        
        dst.mapnum = src.mapnum
        dst.map = src.map
        
        dst.texmapid = src.texmapid
        dst.mapuvid = src.mapuvid
        dst.layernum = src.layernum
        For i = 1 To 4
            dst.layer(i) = src.layer(i)
        Next i
        
    End With
End Sub


'deletes material from LOD
Public Sub BF2DeleteMat(ByRef geom As Long, ByRef lod As Long, ByRef mat As Long)
    MsgBox "Delete geom " & selgeom & " lod " & sellod & " material " & selmat
    
    With vmesh
        If Not .loadok Then Exit Sub
        With .geom(geom).lod(lod)
            
            'check
            If .matnum = 1 Then
                MsgBox "Cannot delete material, geom must have at least one material.", vbExclamation
                Exit Sub
            End If
            
            'shift up
            If mat < .matnum - 1 Then
                Dim i As Long
                For i = mat To .matnum - 2
                    .mat(i) = .mat(i + 1)
                Next i
            End If
            
            'delete
            .matnum = .matnum - 1
            ReDim Preserve .mat(0 To .matnum - 1)
            
        End With
    End With
    
    'sellod = sellod - 1
    frmMain.SelectMesh MakeTag(selgeom, sellod - 1, -1)
    frmMain.FillTreeView
    frmMain.picMain_Paint
End Sub


'deletes LOD from geom
Public Sub BF2DeleteLod(ByRef geom As Long, ByRef lod As Long)
    MsgBox "Delete geom " & selgeom & " lod " & sellod
    
    With vmesh
        If Not .loadok Then Exit Sub
        With .geom(geom)
            
            'check
            If .lodnum = 1 Then
                MsgBox "Cannot delete LOD, geom must have at least one LOD.", vbExclamation
                Exit Sub
            End If
            
            'shift up
            If lod < .lodnum - 1 Then
                Dim i As Long
                For i = lod To .lodnum - 2
                    .lod(i) = .lod(i + 1)
                Next i
            End If
            
            'delete
            .lodnum = .lodnum - 1
            ReDim Preserve .lod(0 To .lodnum - 1)
            
        End With
    End With
    
    'selmat = selmat - 1
    frmMain.SelectMesh MakeTag(selgeom, sellod, selmat - 1)
    frmMain.FillTreeView
    frmMain.picMain_Paint
End Sub


Private Function GetVert(ByVal i As Long) As float3
    With vmesh
        GetVert.x = .vert(i * .xstride + 0)
        GetVert.y = .vert(i * .xstride + 1)
        GetVert.z = .vert(i * .xstride + 2)
    End With
End Function

Private Function GetNorm(ByVal i As Long) As float3
    With vmesh
        GetNorm.x = .vert(i * .xstride + 3 + 0)
        GetNorm.y = .vert(i * .xstride + 3 + 1)
        GetNorm.z = .vert(i * .xstride + 3 + 2)
    End With
End Function

Private Function GetTexc(ByVal i As Long) As float2
    With vmesh 'note: the '7' may not be entirely safe here (also need to make this work for statics)
        GetTexc.x = .vert(i * .xstride + 7 + 0)
        GetTexc.y = .vert(i * .xstride + 7 + 1)
    End With
End Function


Public Sub BF2MatGenTangents(ByRef mat As bf2mat)
    With mat
        'On Error GoTo hell
        On Error Resume Next
        
        'temp tangent array
        'ReDim tan1(0 To vmesh.vertnum - 1) As float3
        ReDim tan2(0 To vmesh.vertnum - 1) As float3
        
        Dim facenum As Long
        facenum = .inum / 3
        
        'compute tangents
        Dim i As Long
        For i = 0 To facenum - 1
            
            Dim i1 As Long
            Dim i2 As Long
            Dim i3 As Long
            i1 = .vstart + vmesh.Index(.istart + (i * 3) + 0)
            i2 = .vstart + vmesh.Index(.istart + (i * 3) + 1)
            i3 = .vstart + vmesh.Index(.istart + (i * 3) + 2)
            
            Dim v1 As float3
            Dim v2 As float3
            Dim v3 As float3
            v1 = GetVert(i1)
            v2 = GetVert(i2)
            v3 = GetVert(i3)
            
            Dim uv1 As float2
            Dim uv2 As float2
            Dim uv3 As float2
            uv1 = GetTexc(i1)
            uv2 = GetTexc(i2)
            uv3 = GetTexc(i3)
            
            Dim x1 As Single
            Dim x2 As Single
            Dim y1 As Single
            Dim y2 As Single
            Dim z1 As Single
            Dim z2 As Single
            x1 = v2.x - v1.x
            x2 = v3.x - v1.x
            y1 = v2.y - v1.y
            y2 = v3.y - v1.y
            z1 = v2.z - v1.z
            z2 = v3.z - v1.z
            
            Dim s1 As Single
            Dim s2 As Single
            s1 = uv2.x - uv1.x
            s2 = uv3.x - uv1.x
            
            Dim t1 As Single
            Dim t2 As Single
            t1 = uv2.y - uv1.y
            t2 = uv3.y - uv1.y
            
            Dim r As Single
            Dim d As Single
            d = (s1 * t2 - s2 * t1)
            If d = 0 Then
                r = 0
            Else
                r = 1 / d
            End If
            
            'Dim sdir As float3
            'sdir.x = (t2 * x1 - t1 * x2) * r
            'sdir.y = (t2 * y1 - t1 * y2) * r
            'sdir.z = (t2 * z1 - t1 * z2) * r
            
            Dim tdir As float3
            tdir.x = (s1 * x2 - s2 * x1) * r
            tdir.y = (s1 * y2 - s2 * y1) * r
            tdir.z = (s1 * z2 - s2 * z1) * r
            
            'tan1(i1) = AddFloat3(tan1(i1), sdir)
            'tan1(i2) = AddFloat3(tan1(i2), sdir)
            'tan1(i3) = AddFloat3(tan1(i3), sdir)
            
            tan2(i1) = AddFloat3(tan2(i1), tdir)
            tan2(i2) = AddFloat3(tan2(i2), tdir)
            tan2(i3) = AddFloat3(tan2(i3), tdir)
            
        Next i
        
        Dim tangoff As Long
        tangoff = BF2MeshGetTangOffset()
        
        'ortho-normalize
        For i = .vstart To .vstart + .vnum - 1
            
            Dim n As float3
            n = GetNorm(i)
            
            Dim t As float3
            't = Normalize(SubFloat3(tan1(i), ScaleFloat3(n, DotProduct(n, tan1(i)))))
            t.x = vmesh.vert(i * vmesh.xstride + tangoff + 0)
            t.y = vmesh.vert(i * vmesh.xstride + tangoff + 1)
            t.z = vmesh.vert(i * vmesh.xstride + tangoff + 2)
            
            vmesh.xtan(i).x = t.x
            vmesh.xtan(i).y = t.y
            vmesh.xtan(i).z = t.z
            
            'vmesh.xtan(i).x = vmesh.vert(i * stride + tangoff + 0)
            'vmesh.xtan(i).y = vmesh.vert(i * stride + tangoff + 1)
            'vmesh.xtan(i).z = vmesh.vert(i * stride + tangoff + 2)
            
            'calculate handedness
            If (DotProduct(CrossProduct(n, t), tan2(i)) > 0) Then
                vmesh.xtan(i).w = 1
            Else
                vmesh.xtan(i).w = -1
            End If
            
        Next i
        
    End With
    
    Exit Sub
hell:
    MsgBox "BF2MatGenTangents" & vbLf & err.description, vbCritical
End Sub


'computes bi-tangents
Public Sub BF2ComputeTangents()
    With vmesh
        If Not .loadok Then Exit Sub
        
        Dim i As Long
        
        Dim stride As Long
        stride = .vertstride / 4
        
        Dim normoff As Long
        Dim tangoff As Long
        normoff = BF2MeshGetNormOffset()
        tangoff = BF2MeshGetTangOffset()
        
        'allocate
        ReDim .xtan(0 To .vertnum - 1)
        'For i = 0 To .vertnum - 1
        '    .xtan(i).w = 0
        'Next i
        
        If 111 = 111 Then
            
            Dim g As Long
            For g = 0 To .geomnum - 1
                Dim L As Long
                For L = 0 To .geom(g).lodnum - 1
                    Dim m As Long
                    For m = 0 To .geom(g).lod(L).matnum - 1
                        BF2MatGenTangents .geom(g).lod(L).mat(m)
                    Next m
                Next L
            Next g
            
        Else
            'determine tangent W by triangle sign
            
            Dim facenum As Long
            facenum = .indexnum / 3
            For i = 0 To facenum - 1
                
                Dim i1 As Long
                Dim i2 As Long
                Dim i3 As Long
                i1 = .Index(i * 3 + 0)
                i2 = .Index(i * 3 + 1)
                i3 = .Index(i * 3 + 2)
                
                Dim uv1 As float2
                Dim uv2 As float2
                Dim uv3 As float2
                uv1 = GetTexc(i1)
                uv2 = GetTexc(i2)
                uv3 = GetTexc(i3)
                
                Dim s As Single
                s = TriangleSign(uv1, uv2, uv3)
               ' .xtan(i1).w = .xtan(i1).w + s
               ' .xtan(i2).w = .xtan(i2).w + s
               ' .xtan(i3).w = .xtan(i3).w + s
                .xtan(i1).w = s
                .xtan(i2).w = s
                .xtan(i3).w = s
                
            Next i
            
            'copy tangs
            For i = 0 To .vertnum - 1
                
                'get tangent
                Dim t As float3
                t.x = vmesh.vert(i * stride + tangoff + 0)
                t.y = vmesh.vert(i * stride + tangoff + 1)
                t.z = vmesh.vert(i * stride + tangoff + 2)
                
                'normalize
               ' Dim w As Single
               ' w = .xtan(i).w
               ' If w < 0 Then w = -1
               ' If w > 0 Then w = 1
               ' If w = 0 Then w = 1
                
                'store
                .xtan(i).x = t.x
                .xtan(i).y = t.y
                .xtan(i).z = t.z
               ' .xtan(i).w = w
                
            Next i
        End If
        
    End With
End Sub


'returns sign of triangle
Private Function TriangleSign(ByRef v1 As float2, ByRef v2 As float2, ByRef v3 As float2) As Single
    If ((v2.y - v1.y) - (v2.x - v1.x)) + _
       ((v3.y - v2.y) - (v3.x - v2.x)) + _
       ((v1.y - v3.y) - (v1.x - v3.x)) > 0 Then
        TriangleSign = 1
    Else
        TriangleSign = -1
    End If
End Function
