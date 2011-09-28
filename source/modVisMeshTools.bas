Attribute VB_Name = "BF2_MeshTools"
Option Explicit


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
Private Sub BF2VerifyMat(ByRef mat As bf2_mat)
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


Public Sub PasteMaterial(ByRef dst As bf2_mat, ByRef src As bf2_mat)
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


