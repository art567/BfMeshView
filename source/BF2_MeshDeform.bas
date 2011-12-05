Attribute VB_Name = "BF2_MeshDeform"
Option Explicit


'deforms skinnedmesh by skeleton
Public Sub BF2MeshDeformSM()
    With vmesh
        On Error GoTo errhandler
        
        If Not .loadok Then Exit Sub
        If Not .drawok Then Exit Sub
        If Not bf2ske.loaded Then Exit Sub
        
        Dim stride As Long
        stride = vmesh.vertstride / 4
        
        Dim woff As Long
        'woff = BF2MeshGetWeightOffset
        woff = 6
        
        If Not .hasSkinVerts Then
            .hasSkinVerts = True
            ReDim .skinvert(0 To .vertnum - 1)
            ReDim .skinnorm(0 To .vertnum - 1)
        End If
        
        Dim i As Long
        Dim j As Long
        Dim m As Long
        Dim v As float3
        Dim n As float3
        Dim vw As bf2skinweight
        
        ''reset all vertices
        'For i = 0 To .vertnum - 1
        '    .skinvert(i).x = .vert(i * stride + 0)
        '    .skinvert(i).y = .vert(i * stride + 1)
        '    .skinvert(i).z = .vert(i * stride + 2)
        'Next i
        
        With .geom(selgeom).lod(sellod)
            
            'can't deform if geom lod has no rigs
            If .rignum = 0 Then Exit Sub
            
            'fill per-bone skinning matrices
            For i = 0 To .rignum - 1
                With .rig(i)
                    For j = 0 To .bonenum - 1
                        With .bone(j)
                            If .id > -1 And .id < bf2ske.nodenum Then
                                .skinmat = mat4mult(.matrix, bf2ske.node(.id).worldmat)
                            Else
                                mat4identity .skinmat
                            End If
                        End With
                    Next j
                End With
            Next i
            
            'deform vertices
            For m = 0 To .matnum - 1
                For i = .mat(m).vstart To .mat(m).vstart + .mat(m).vnum - 1
                    If .rig(m).bonenum > 0 Then
                        
                        'get vertex position
                        v.X = vmesh.vert(i * stride + 0)
                        v.y = vmesh.vert(i * stride + 1)
                        v.z = vmesh.vert(i * stride + 2)
                        
                        'get normal
                        n.X = vmesh.vert(i * stride + 3 + 0)
                        n.y = vmesh.vert(i * stride + 3 + 1)
                        n.z = vmesh.vert(i * stride + 3 + 2)
                        
                        'get vertex weight
                        CopyMem VarPtr(vw), VarPtr(vmesh.vert(i * stride + woff)), 8
                        
                        'deform vertex
                        Dim tv As float3 'temp vert
                        Dim tn As float3 'temp normal
                        Dim dv As float3 'deformed vert
                        Dim dn As float3 'deformed normal
                        
                        dv = float3(0, 0, 0)
                        dn = float3(0, 0, 0)
                            
                        'bone 1
                        tv = mat4transvec(.rig(m).bone(vw.b1).skinmat, v)
                        tn = mat4rotvec(.rig(m).bone(vw.b1).skinmat, n)
                        dv.X = dv.X + tv.X * vw.w
                        dv.y = dv.y + tv.y * vw.w
                        dv.z = dv.z + tv.z * vw.w
                        
                        dn.X = dn.X + tn.X * vw.w
                        dn.y = dn.y + tn.y * vw.w
                        dn.z = dn.z + tn.z * vw.w
                        
                        'bone 2
                        tv = mat4transvec(.rig(m).bone(vw.b2).skinmat, v)
                        tn = mat4rotvec(.rig(m).bone(vw.b2).skinmat, n)
                        dv.X = dv.X + tv.X * (1 - vw.w)
                        dv.y = dv.y + tv.y * (1 - vw.w)
                        dv.z = dv.z + tv.z * (1 - vw.w)
                        
                        dn.X = dn.X + tn.X * (1 - vw.w)
                        dn.y = dn.y + tn.y * (1 - vw.w)
                        dn.z = dn.z + tn.z * (1 - vw.w)
                        
                        'store deformed attributes
                        vmesh.skinvert(i) = dv
                        vmesh.skinnorm(i) = dn
                        
                    Else
                        vmesh.skinvert(i).X = vmesh.vert(i * stride + 0)
                        vmesh.skinvert(i).y = vmesh.vert(i * stride + 1)
                        vmesh.skinvert(i).z = vmesh.vert(i * stride + 2)
                        
                        vmesh.skinnorm(i).X = vmesh.vert(i * stride + 3 + 0)
                        vmesh.skinnorm(i).y = vmesh.vert(i * stride + 3 + 1)
                        vmesh.skinnorm(i).z = vmesh.vert(i * stride + 3 + 2)
                    End If
                Next i
            Next m
        End With
        
    End With
    
    Exit Sub
errhandler:
    MsgBox "BF2MeshDeform" & vbLf & err.description, vbCritical
End Sub


'deforms bundledmesh with CON nodes
Public Sub BF2MeshDeformBM()
    Dim i As Long

    'reset
    nodetransformnum = 40
    For i = 0 To 40 - 1
        mat4identity nodetransform(i)
    Next i
    
    With bf2con
        If Not .loaded Then Exit Sub
        If .nodenum <= 1 Then Exit Sub
        If .partnum <= 1 Then Exit Sub
    End With
    With vmesh
        If Not .loadok Then Exit Sub
        If Not .drawok Then Exit Sub
        
        If Not .hasSkinVerts Then
            .hasSkinVerts = True
            ReDim .skinvert(0 To .vertnum - 1)
            ReDim .skinnorm(0 To .vertnum - 1)
        End If
        
        For i = 0 To .geomnum - 1
            With .geom(i)
                Dim j As Long
                For j = 0 To .lodnum - 1
                    BF2MeshDeformGeomLod i, j
                Next j
            End With
        Next i
    End With
    
    'fill nodetransform table
    With bf2con
        If .partnum > 0 Then
            nodetransformnum = .partnum
            For i = 0 To .partnum - 1
                nodetransform(i) = .node(.part(i)).wtrans
            Next i
        End If
    End With
    
End Sub

Public Sub BF2MeshDeformGeomLod(ByRef geom As Long, ByRef lod As Long)
    With vmesh
        On Error GoTo errhandler
        
        Dim stride As Long
        stride = vmesh.vertstride / 4
        
        Dim woff As Long
        'woff = BF2MeshGetWeightOffset
        woff = 6
        
        Dim i As Long
        Dim j As Long
        Dim m As Long
        Dim v As float3
        Dim n As float3
        Dim vw As bf2vw
        
        ''reset all vertices
        'For i = 0 To .vertnum - 1
        '    .skinvert(i).x = .vert(i * stride + 0)
        '    .skinvert(i).y = .vert(i * stride + 1)
        '    .skinvert(i).z = .vert(i * stride + 2)
        'Next i
        
        With .geom(geom).lod(lod)
            
            'deform vertices
            For m = 0 To .matnum - 1
                For i = .mat(m).vstart To .mat(m).vstart + .mat(m).vnum - 1
                    
                    'get vertex position
                    v.X = vmesh.vert(i * stride + 0)
                    v.y = vmesh.vert(i * stride + 1)
                    v.z = vmesh.vert(i * stride + 2)
                    
                    'get normal
                    n.X = vmesh.vert(i * stride + 3 + 0)
                    n.y = vmesh.vert(i * stride + 3 + 1)
                    n.z = vmesh.vert(i * stride + 3 + 2)
                    
                    'get vertex weight
                    CopyMem VarPtr(vw), VarPtr(vmesh.vert(i * stride + woff)), 4
                    
                    Dim p As Long
                    p = vw.b1
                    If p < 0 Then p = 0
                    If p > bf2con.partnum - 1 Then p = 0
                    
                    Dim nodeid As Long
                    nodeid = bf2con.part(p)
                    
                    'deform vertex
                    vmesh.skinvert(i) = mat4transvec(bf2con.node(nodeid).wtrans, v)
                    vmesh.skinnorm(i) = mat4rotvec(bf2con.node(nodeid).wtrans, n)
                    
                Next i
            Next m
        End With
        
    End With
    
    Exit Sub
errhandler:
    MsgBox "BF2MeshDeform3" & vbLf & err.description, vbCritical
End Sub

