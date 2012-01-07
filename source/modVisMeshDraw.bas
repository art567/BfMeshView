Attribute VB_Name = "BF2_MeshDraw"
Option Explicit

Public Enum dw_drawmode
    dm_normal
    dm_vertorder
    dm_overdraw
End Enum

Public draw_mode As Long
Public bakemode As Long

Public selgeom As Long
Public selsub As Long
Public sellod As Long
Public selmat As Long
Public seltex As Long
Public seldefault As Long

Private lmUVptr As Long 'current lightmap UV pointer


'draws mesh
Public Sub DrawVisMesh()
    If Not vmesh.drawok Then Exit Sub
    On Error GoTo errorhandler
    
    With vmesh
        If Not .loadok Then Exit Sub
            
        'deform mesh
        If bf2baf.loaded Then
            BF2MeshDeformSM
        End If
        
        'draw mesh
        Select Case draw_mode
        Case dm_normal
            DrawVisMeshLod .geom(selgeom).lod(sellod)
        Case dm_vertorder
            DrawVisMeshIndexColors .geom(selgeom).lod(sellod)
        Case dm_overdraw
            DrawVisMeshOverdraw .geom(selgeom).lod(sellod)
        End Select
        
    End With
    
    'succes
    Exit Sub
errorhandler:
    
    If err.Number = 11 Then
        'ignore, Intel I5 FPU bug it seems
        Exit Sub
    End If
    
    vmesh.drawok = False
    MsgBox "DrawVisMesh" & vbLf & err.description, vbCritical
End Sub


'draws mesh
Public Sub DrawVisMeshLod(ByRef mesh As bf2lod)
Dim i As Long
Dim j As Long
Dim vptr As Long 'vertex array pointer
Dim tptr As Long 'texcoord array pointer
Dim nptr As Long 'normal array pointer
Dim iptr As Long 'index array pointer
Dim stride As Long
Dim texchans As Long
    
    'array start pointers
    With vmesh 'TODO: use vertex attribute table!
        Select Case .vertstride
        Case 48
            vptr = VarPtr(.vert(0))
            nptr = VarPtr(.vert(3))
            tptr = VarPtr(.vert(7))
            texchans = 1
        Case 52
            vptr = VarPtr(.vert(0))
            nptr = VarPtr(.vert(3))
            tptr = VarPtr(.vert(8))
            texchans = 1
        Case 56
            vptr = VarPtr(.vert(0))
            nptr = VarPtr(.vert(3))
            tptr = VarPtr(.vert(7))
            texchans = 1
        Case 72
            vptr = VarPtr(.vert(0))
            nptr = VarPtr(.vert(3))
            tptr = VarPtr(.vert(7))
            texchans = 1
        Case 80
            vptr = VarPtr(.vert(0))
            nptr = VarPtr(.vert(3))
            tptr = VarPtr(.vert(7))
            texchans = 4
        Case Else
            'this usually works
            If .vertstride >= 12 Then vptr = VarPtr(.vert(0))
            If .vertstride >= 24 Then nptr = VarPtr(.vert(3))
            If .vertstride >= 36 Then tptr = VarPtr(.vert(7))
            texchans = 1
        End Select
        iptr = VarPtr(.Index(0))
        stride = .vertstride / 4
    End With
    
    With mesh
        For i = 0 To .matnum - 1
            With .mat(i)
                
                Dim vptroff As Long
                Dim nptroff As Long
                Dim tptroff As Long
                Dim iptroff As Long
                If vptr Then vptroff = vptr + (.vstart * vmesh.vertstride)
                If nptr Then nptroff = nptr + (.vstart * vmesh.vertstride)
                If tptr Then tptroff = tptr + (.vstart * vmesh.vertstride)
                If iptr Then iptroff = iptr + (.istart * 2) 'sizeof(uint16)
                
                lmUVptr = tptroff + (8 * (vmesh.uvnum - 1))
                
                If vmesh.hasSkinVerts Then
                    vptroff = VarPtr(vmesh.skinvert(0)) + (.vstart * 12)
                    nptroff = VarPtr(vmesh.skinnorm(0)) + (.vstart * 12)
                    
                    'fill nodetransforms
                    If vmesh.isSkinnedMesh And .glslprog > 0 Then
                        nodetransformnum = mesh.rig(i).bonenum
                        For j = 0 To mesh.rig(i).bonenum - 1
                            nodetransform(j) = mesh.rig(i).bone(j).skinmat
                        Next j
                    End If
                End If
                
                Dim vcount As Long
                Dim icount As Long
                vcount = .vnum
                icount = .inum
                
                'draw polygons
                If view_poly Then
                    
                    Dim showOnlyTexture As Boolean
                    showOnlyTexture = (seltex > -1 And selmat = i)
                    
                    If .glslprog > 0 And Not showOnlyTexture Then
                        'GLSL shader pipeline
                        
                        If view_edges Or view_verts Then
                            glPolygonOffset 1, 1
                            glEnable GL_POLYGON_OFFSET_FILL
                        End If
                        
                        'prepare
                        glUseProgram .glslprog
                        SetUniforms .glslprog, mesh.mat(i)
                        
                        'alpha mode
                        Select Case .alphamode
                        Case 1:
                            glBlendFunc GL_SRC_ALPHA, GL_ONE_MINUS_SRC_ALPHA
                            glDepthMask GL_FALSE
                            glEnable GL_BLEND
                        Case 2:
                            glAlphaFunc GL_GREATER, 0.5
                            glEnable GL_ALPHA_TEST
                        End Select
                        If .alphaTest > 0 Then
                            glAlphaFunc GL_GREATER, .alphaTest
                            glEnable GL_ALPHA_TEST
                        End If
                        If .twosided Then
                            glDisable GL_CULL_FACE
                        End If
                        
                        'bind textures
                        For j = 0 To .mapnum - 1
                            glActiveTexture GL_TEXTURE0 + j
                            glClientActiveTexture GL_TEXTURE0 + j
                            If .texmapid(j) > 0 Then
                                glBindTexture GL_TEXTURE_2D, texmap(.texmapid(j)).tex
                            Else
                                glBindTexture GL_TEXTURE_2D, dummytex
                            End If
                            glBindTexture GL_TEXTURE_CUBE_MAP, 0
                        Next j
                        If .hasEnvMap Then
                            glActiveTexture GL_TEXTURE0 + envmapChan
                            glClientActiveTexture GL_TEXTURE0 + envmapChan
                            glBindTexture GL_TEXTURE_CUBE_MAP, envmapTex
                        End If
                        
                        'draw
                        drawfacesX .vstart, .vnum, .istart, .inum
                        
                        'reset
                        glUseProgram 0
                        glActiveTexture GL_TEXTURE0
                        glDisable GL_ALPHA_TEST
                        glDisable GL_BLEND
                        glEnable GL_CULL_FACE
                        glDepthMask GL_TRUE
                        
                        If view_edges Or view_verts Then
                            glDisable GL_POLYGON_OFFSET_FILL
                        End If
                    Else
                        'fixed function pipeline
                        
                        glActiveTexture GL_TEXTURE0
                        glClientActiveTexture GL_TEXTURE0
                        
                        'prepare stuff
                        If view_lighting Then
                            glEnable GL_LIGHTING
                        End If
                        If view_edges Or view_verts Then
                            glPolygonOffset 1, 1
                            glEnable GL_POLYGON_OFFSET_FILL
                        End If
                        
                        'draw geometry
                        Dim texcoff As Long
                        If showOnlyTexture Then
                                
                            'render single unlit pass with texture
                            
                            If view_lighting Then glDisable GL_LIGHTING
                            glBindTexture GL_TEXTURE_2D, texmap(.texmapid(seltex)).tex
                            glEnable GL_TEXTURE_2D
                            glColor3f 1, 1, 1
                            
                            'determine the UV channel index for this texture map
                            texcoff = .mapuvid(seltex)
                            
                            'draw geometry
                            drawfaces vptroff, nptroff, tptroff + (8 * texcoff), iptroff, icount
                            
                            'reset stuff
                            glDisable GL_TEXTURE_2D
                            If view_lighting Then glEnable GL_LIGHTING
                        Else
                            If view_textures And .layernum > 0 Then
                                
                                'render each texture layer as seperate pass
                                For j = 1 To .layernum
                                    
                                    'get texture for this layer
                                    Dim texmapid As Long
                                    texmapid = .layer(j).texmapid
                                    
                                    If texmapid > 0 Then
                                        
                                        texcoff = .layer(j).texcoff
                                        
                                        If .layer(j).blend Then
                                            glBlendFunc .layer(j).blendsrc, .layer(j).blenddst
                                            glEnable GL_BLEND
                                        End If
                                        If .layer(j).alphaTest Then
                                            glEnable GL_ALPHA_TEST
                                            glAlphaFunc GL_GREATER, .layer(j).alpharef
                                        End If
                                        If .layer(j).twosided Then
                                            glDisable GL_CULL_FACE
                                        End If
                                        If .layer(j).lighting And view_lighting Then
                                            glEnable GL_LIGHTING
                                        End If
                                        glDepthMask .layer(j).depthWrite
                                        glDepthFunc .layer(j).depthfunc
                                        
                                        glBindTexture GL_TEXTURE_2D, texmap(texmapid).tex
                                        glEnable GL_TEXTURE_2D
                                        glColor4f 1, 1, 1, 1
                                        
                                        drawfaces vptroff, nptroff, tptroff + (8 * texcoff), iptroff, icount
                                        
                                        glDepthMask True
                                        glDisable GL_BLEND
                                        glDisable GL_TEXTURE_2D
                                        glDisable GL_ALPHA_TEST
                                        glEnable GL_CULL_FACE
                                        glDepthFunc GL_LESS
                                        glDisable GL_LIGHTING
                                    Else
                                        glColor3f 0.75, 0.75, 0.75
                                        drawfaces vptroff, nptroff, 0, iptroff, icount
                                    End If
                                Next j
                            Else
                                glColor3f 0.75, 0.75, 0.75
                                drawfaces vptroff, nptroff, 0, iptroff, icount
                            End If
                        End If
                        
                        'reset stuff
                        If view_edges Or view_verts Then
                            glDisable GL_POLYGON_OFFSET_FILL
                        End If
                        If view_lighting Then
                            glDisable GL_LIGHTING
                        End If
                        If view_textures Then
                            glDisable GL_TEXTURE_2D
                        End If
                    End If
                    
                    'draw edges
                    If view_edges And Not view_wire Then
                        glColor4f 1, 1, 1, 0.1
                        StartAALine 1.3
                        glPolygonMode GL_FRONT_AND_BACK, GL_LINE
                        drawfaces vptroff, 0, 0, iptroff, icount
                        glPolygonMode GL_FRONT_AND_BACK, GL_FILL
                        EndAALine
                    End If
                    
                End If
            End With 'end material
            
            'draw vertices
            If view_verts Then
                glColor4f 1, 1, 1, 1
                If vmesh.hasSkinVerts Then
                    glVertexPointer 3, GL_FLOAT, 0, ByVal vptroff
                Else
                    glVertexPointer 3, GL_FLOAT, vmesh.vertstride, ByVal vptroff
                End If
                StartAAPoint 4
                glEnableClientState GL_VERTEX_ARRAY
                glDrawArrays GL_POINTS, 0, vcount
                glDisableClientState GL_VERTEX_ARRAY
                EndAAPoint
            End If
            
            'draw selected UV verts
            Dim showvertsel As Boolean
            If uveditor_isloaded Then
                If frmUvEdit.Visible Then
                    frmUvEdit.SetVertFlags
                    showvertsel = True
                End If
            End If
            If frmTransform.Visible Then
                showvertsel = True
            End If
            If toolmode = 1 Then
                showvertsel = True
            End If
            If showvertsel Then
                glColor4f 1, 0, 0, 1
                StartAAPoint 5
                glBegin GL_POINTS
                    For j = 0 To vmesh.vertnum - 1
                        If vmesh.vertflag(j) Then
                            If vmesh.vertsel(j) Then
                                If vmesh.hasSkinVerts Then
                                    glVertex3fv vmesh.skinvert(j).x
                                Else
                                    glVertex3fv vmesh.vert(j * stride)
                                End If
                            End If
                        End If
                    Next j
                glEnd
                EndAAPoint
            End If
            
            Const s As Single = 0.05 'normal/tangent scale
            Dim k As Long
            Dim n As float3
            Dim vi As Long 'vertex index
            Dim ni As Long 'normal index
            
            'draw tangents
            If view_tangents Then
                Dim q As Long
                Dim v As float3     'vertex
                Dim t As float4     'tangent vector
                Dim b As float3     'binormal vector
                
                Dim normoff As Long
                Dim tangoff As Long
                normoff = 3
                tangoff = ((vmesh.vertstride - 24) / 4) + 3
                'tangoff = ((vmesh.vertstride - 24) + 12) / 4 'todo: try this (just for code clarity)
                
                StartAALine 1.3
                glBegin GL_LINES
                    For j = 0 To .matnum - 1
                        For k = 0 To .mat(j).vnum - 1
                            vi = ((.mat(j).vstart + k) * stride)
                            
                            'get vertex
                            v.x = vmesh.vert(vi + 0)
                            v.y = vmesh.vert(vi + 1)
                            v.z = vmesh.vert(vi + 2)
                            
                            'get normal
                            n.x = vmesh.vert(vi + normoff + 0)
                            n.y = vmesh.vert(vi + normoff + 1)
                            n.z = vmesh.vert(vi + normoff + 2)
                            
                            'get tangent
                            Dim ot As float3
                            ot.x = vmesh.vert(vi + tangoff + 0)
                            ot.y = vmesh.vert(vi + tangoff + 1)
                            ot.z = vmesh.vert(vi + tangoff + 2)
                            t = vmesh.xtan(.mat(j).vstart + k)
                            
                            'get binormal
                            b = CrossProduct(n, float3(t.x, t.y, t.z))
                            
                            'rescale
                            t.x = v.x + t.x * s
                            t.y = v.y + t.y * s
                            t.z = v.z + t.z * s
                            b.x = v.x + b.x * s * t.w
                            b.y = v.y + b.y * s * t.w
                            b.z = v.z + b.z * s * t.w
                            
                            'draw tangent
                            glColor4f 1, 0.5, 0.5, 0.5
                            glVertex3fv v.x
                            glVertex3fv t.x
                            
                            'draw bitangent
                            glColor4f 0.5, 1, 0.5, 0.5
                            glVertex3fv v.x
                            glVertex3fv b.x
                            
                            'draw bitangent
                            glColor4f 0, 0, 0, 0.5 '1, 0.5, 1, 0.5
                            glVertex3fv v.x
                            glVertex3f v.x + ot.x * s * 0.9, v.y + ot.y * s * 0.9, v.z + ot.z * s * 0.9
                            
                        Next k
                    Next j
                glEnd
                EndAALine
            End If
            
            'draw normals (note: we draw these last since they tend to overdraw tangents without depth testing)
            If view_normals Then
                glColor4f 0, 1, 1, 0.5
                StartAALine 1.3
                
                stride = vmesh.vertstride / 4
                glBegin GL_LINES
                    For j = 0 To .matnum - 1
                        For k = 0 To .mat(j).vnum - 1
                            vi = ((.mat(j).vstart + k) * stride) + 0
                            ni = ((.mat(j).vstart + k) * stride) + 3
                            
                            n.x = vmesh.vert(vi + 0) + vmesh.vert(ni + 0) * s
                            n.y = vmesh.vert(vi + 1) + vmesh.vert(ni + 1) * s
                            n.z = vmesh.vert(vi + 2) + vmesh.vert(ni + 2) * s
                            
                            glVertex3fv vmesh.vert(vi)
                            glVertex3fv n.x
                            
                        Next k
                    Next j
                glEnd
                
                EndAALine
            End If
            
        Next i
        
        'draw skin matrices
        If view_bonesys And vmesh.hasSkinVerts = False Then
            
            Dim im As matrix4
            
            'draw bones matrices
            glDisable GL_DEPTH_TEST
            StartAAPoint 9
            StartAALine 1.3
            For i = 0 To .rignum - 1
                For j = 0 To .rig(i).bonenum - 1
                    glPushMatrix
                        'If vmesh.hasSkinVerts Then
                        '    glMultMatrixf .rig(i).bone(j).skinmat.m(0)
                        'Else
                            GetInverseMat4 .rig(i).bone(j).matrix.m, im.m
                            glMultMatrixf im.m(0)
                        'End If
                        
                        glBegin GL_POINTS
                            glColor3f 1, 1, 0
                            glVertex3f 0, 0, 0
                        glEnd
                        DrawPivot 0.025
                    glPopMatrix
                Next j
            Next i
            EndAALine
            EndAAPoint
            glEnable GL_DEPTH_TEST
            
        End If
        
        'draw bounds
        If view_bounds Then
            
            'mesh bounds
            StartAALine 1.3
            glColor3f 1, 1, 0
            DrawBox .min, .max
            EndAALine
            
            'per material/drawcall bounds
            If vmesh.head.version = 11 Then
                
                glLineStipple 1, &HF0F
                glEnable GL_LINE_STIPPLE
                StartAALine 1.3
                glColor3f 1, 0.5, 0
                For i = 0 To .matnum - 1
                    DrawBox .mat(i).mmin, .mat(i).mmax
                Next i
                EndAALine
                glDisable GL_LINE_STIPPLE
                
            End If
            
        End If
                
    End With
End Sub


'draws material group
Private Sub drawfaces(ByVal vptr As Long, ByVal nptr As Long, ByVal tptr As Long, ByVal iptr As Long, ByVal inum As Long)
    With vmesh
        
        'no vertex stride if skin is deformed
        Dim vs As Long
        If vmesh.hasSkinVerts Then
            vs = 0
        Else
            vs = .vertstride
        End If
        
        'substitute vertex coordinates for texcoords when baking
        Dim vertFrags As Long
        Dim vrealptr As Long
        If bakemode Then
            vertFrags = 2
            vrealptr = lmUVptr
            glDisable GL_LIGHTING
            glDisable GL_CULL_FACE
            glDisable GL_DEPTH_TEST
        Else
            vertFrags = 3
            vrealptr = vptr
        End If
        
        If vptr Then glVertexPointer vertFrags, GL_FLOAT, vs, ByVal vrealptr
        If nptr Then glNormalPointer GL_FLOAT, vs, ByVal nptr
        If tptr Then glTexCoordPointer 2, GL_FLOAT, .vertstride, ByVal tptr
        
        If vptr Then glEnableClientState GL_VERTEX_ARRAY
        If nptr Then glEnableClientState GL_NORMAL_ARRAY
        If tptr Then glEnableClientState GL_TEXTURE_COORD_ARRAY
        
        glDrawElements GL_TRIANGLES, inum, GL_UNSIGNED_SHORT, ByVal iptr
        
        If vptr Then glDisableClientState GL_VERTEX_ARRAY
        If nptr Then glDisableClientState GL_NORMAL_ARRAY
        If tptr Then glDisableClientState GL_TEXTURE_COORD_ARRAY
        
        If bakemode Then
            glEnable GL_DEPTH_TEST
        End If
        
    End With
End Sub

'draws material faces
Private Sub drawfacesX(ByVal voff As Long, vnum As Long, ByVal ioff As Long, ByVal inum As Long)
    With vmesh
        
        Dim vptr As Long
        vptr = VarPtr(.vert(0)) + voff * .vertstride
        
        Dim iptr As Long
        iptr = VarPtr(.Index(0)) + ioff * 2
        
        'enable vertex attributes
        Dim i As Long
        For i = 0 To .vertattribnum - 1
            Dim off As Long
            off = vptr + .vertattrib(i).offset
            
            Select Case .vertattrib(i).usage
            Case 0: 'position
                glEnableClientState GL_VERTEX_ARRAY
                glVertexPointer 3, GL_FLOAT, .vertstride, ByVal off
            Case 1: 'blend weight
                glClientActiveTexture GL_TEXTURE1
                glEnableClientState GL_TEXTURE_COORD_ARRAY
                glTexCoordPointer 1, GL_FLOAT, .vertstride, ByVal off
            Case 2: 'blend indices
                glEnableClientState GL_COLOR_ARRAY
                glColorPointer 4, GL_UNSIGNED_BYTE, .vertstride, ByVal off
            Case 3: 'normal
                glEnableClientState GL_NORMAL_ARRAY
                glNormalPointer GL_FLOAT, .vertstride, ByVal off
            Case 5: 'uv1
                glClientActiveTexture GL_TEXTURE0
                glEnableClientState GL_TEXTURE_COORD_ARRAY
                glTexCoordPointer 2, GL_FLOAT, .vertstride, ByVal off
            Case 6: 'tangent
                glClientActiveTexture GL_TEXTURE5
                glEnableClientState GL_TEXTURE_COORD_ARRAY
                'glTexCoordPointer 3, GL_FLOAT, .vertstride, ByVal off
                glTexCoordPointer 4, GL_FLOAT, 0, ByVal VarPtr(.xtan(voff))
            Case 261: 'uv2
                glClientActiveTexture GL_TEXTURE1
                glEnableClientState GL_TEXTURE_COORD_ARRAY
                glTexCoordPointer 2, GL_FLOAT, .vertstride, ByVal off
            Case 517: 'uv3
                glClientActiveTexture GL_TEXTURE2
                glEnableClientState GL_TEXTURE_COORD_ARRAY
                glTexCoordPointer 2, GL_FLOAT, .vertstride, ByVal off
            Case 773: 'uv4
                glClientActiveTexture GL_TEXTURE3
                glEnableClientState GL_TEXTURE_COORD_ARRAY
                glTexCoordPointer 2, GL_FLOAT, .vertstride, ByVal off
            Case 1029: 'uv5
                glClientActiveTexture GL_TEXTURE4
                glEnableClientState GL_TEXTURE_COORD_ARRAY
                glTexCoordPointer 2, GL_FLOAT, .vertstride, ByVal off
            End Select
        Next i
        
        'draw triangles
        glDrawElements GL_TRIANGLES, inum, GL_UNSIGNED_SHORT, ByVal iptr
        
        'disable vertex attributes
        For i = 0 To .vertattribnum - 1
            
            Select Case .vertattrib(i).usage
            Case 0: 'position
                glDisableClientState GL_VERTEX_ARRAY
            Case 1: 'blend weight
                glClientActiveTexture GL_TEXTURE1
                glDisableClientState GL_TEXTURE_COORD_ARRAY
            Case 2: 'blend indices
                glDisableClientState GL_COLOR_ARRAY
            Case 3: 'normal
                glDisableClientState GL_NORMAL_ARRAY
            Case 5: 'uv1
                glClientActiveTexture GL_TEXTURE0
                glDisableClientState GL_TEXTURE_COORD_ARRAY
            Case 6: 'tangent
                glClientActiveTexture GL_TEXTURE5
                glDisableClientState GL_TEXTURE_COORD_ARRAY
            Case 261: 'uv2
                glClientActiveTexture GL_TEXTURE1
                glDisableClientState GL_TEXTURE_COORD_ARRAY
            Case 517: 'uv3
                glClientActiveTexture GL_TEXTURE2
                glDisableClientState GL_TEXTURE_COORD_ARRAY
            Case 773: 'uv4
                glClientActiveTexture GL_TEXTURE3
                glDisableClientState GL_TEXTURE_COORD_ARRAY
            Case 1029: 'uv5
                glClientActiveTexture GL_TEXTURE4
                glDisableClientState GL_TEXTURE_COORD_ARRAY
            End Select
        Next i
        
    End With
End Sub


'draws triangles with index color
Private Sub DrawVisMeshIndexColors(ByRef lod As bf2lod)
Dim m As Long
Dim i As Long
Dim v1 As Long
Dim v2 As Long
Dim v3 As Long
Dim w As Single
Dim c As color4f
Dim ci As Long
Dim stride As Long
    
    glDisable GL_LIGHTING
    glDisable GL_TEXTURE_2D
    
    stride = vmesh.vertstride / 4
    
    With lod
        For m = 0 To .matnum - 1
            
            With .mat(m)
                glBegin GL_TRIANGLES
                For i = 0 To .inum - 1
                    
                    w = i / .inum
                    c.r = colortable(ci).r * w
                    c.g = colortable(ci).g * w
                    c.b = colortable(ci).b * w
                    glColor4fv c.r
                    
                    v1 = (.vstart + vmesh.Index(.istart + i))
                    
                    glVertex3fv vmesh.vert(v1 * stride)
                    
                Next i
                glEnd
            End With
            
            'pick next random color
            ci = ci + 1
            If ci = maxcolors Then
                ci = 0
            End If
            
        Next m
    End With
End Sub


'draws LOD as overdraw mode
Private Sub DrawVisMeshOverdraw(ByRef lod As bf2lod)
Dim m As Long
Dim i As Long
Dim v1 As Long
Dim v2 As Long
Dim v3 As Long
Dim c As color4f
Dim ci As Long
Dim stride As Long
    stride = vmesh.vertstride / 4
    
    glClearStencil 0
    glClear GL_STENCIL_BUFFER_BIT
    glDisable GL_STENCIL_TEST
    
    glEnable GL_DEPTH_TEST
    glDisable GL_LIGHTING
    glDisable GL_BLEND
    glDisable GL_ALPHA_TEST
    glDisable GL_TEXTURE_2D
    
    glEnable GL_STENCIL_TEST
    glStencilFunc GL_ALWAYS, 0, 0
    glStencilOp GL_KEEP, GL_KEEP, GL_INCR
    glColorMask GL_FALSE, GL_FALSE, GL_FALSE, GL_FALSE
    glColor3f 1, 1, 1
    DrawPassX lod
    
    '--- 2d ------------------------------------------------
    
    glMatrixMode GL_PROJECTION
    glPushMatrix
    glLoadIdentity
    
    glMatrixMode GL_MODELVIEW
    glPushMatrix
    glLoadIdentity
    
    glDisable GL_LIGHTING
    glDisable GL_DEPTH_TEST
    glFrontFace GL_CCW
    
    glColor3f 1, 1, 1
    glStencilOp GL_KEEP, GL_KEEP, GL_KEEP
    
    glColorMask GL_TRUE, GL_TRUE, GL_TRUE, GL_TRUE
    
    Dim r As Single
    Dim g As Single
    Dim b As Single
    
    For i = 1 To 10
        glStencilFunc GL_EQUAL, i, &HFFFFFFFF
        
        ColorRamp (i - 1) / 10, r, g, b
        glColor3f r, g, b
        
        glRectf -1, -1, 1, 1
    Next i
    
    glStencilFunc GL_LESS, 10, &HFFFFFFFF
    glColor3f 1, 0, 0
    glRectf -1, -1, 1, 1
    
    glColorMask GL_TRUE, GL_TRUE, GL_TRUE, GL_TRUE
    glDisable GL_STENCIL_TEST
    glEnable GL_DEPTH_TEST
    
    glFrontFace GL_CW
    
    glMatrixMode GL_PROJECTION
    glPopMatrix
    glMatrixMode GL_MODELVIEW
    glPopMatrix
    
End Sub


'...
Private Sub DrawPassX(ByRef lod As bf2lod)
Dim m As Long
Dim i As Long
Dim v1 As Long
Dim stride As Long
    With lod
        stride = vmesh.vertstride / 4
        glBegin GL_TRIANGLES
        For m = 0 To .matnum - 1
            With .mat(m)
                For i = 0 To .inum - 1
                    v1 = (.vstart + vmesh.Index(.istart + i))
                    glVertex3fv vmesh.vert(v1 * stride)
                Next i
            End With
        Next m
        glEnd
    End With
End Sub


'outputs non-gamma corrected 'color rainbow'
Public Sub ColorRamp(ByVal v As Single, ByRef r As Single, ByRef g As Single, ByRef b As Single)
    If v < 0 Then v = 0
    If v > 1 Then v = 1
    If v < 0.25 Then
        r = 0
        g = 4 * v
        b = 1
    ElseIf v < 0.5 Then
        r = 0
        g = 1
        b = 1 + 4 * (0.25 - v)
    ElseIf v < 0.75 Then
        r = 4 * (v - 0.5)
        g = 1
        b = 0
    Else
        r = 1
        g = 1 + 4 * (0.75 - v)
        b = 0
    End If
End Sub

