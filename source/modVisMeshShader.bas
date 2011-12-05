Attribute VB_Name = "BF2_MeshShader"
Option Explicit

Private Sub SetBase(ByRef mat As bf2mat, ByVal i As Long)
    mat.layer(i).texcoff = 0
    mat.layer(i).texmapid = mat.texmapid(0)
    mat.layer(i).depthfunc = GL_LESS
    mat.layer(i).depthWrite = GL_TRUE
    mat.layer(i).lighting = False
    mat.layer(i).blend = False
    mat.layer(i).alphatest = False
    
    Select Case mat.alphamode
    Case 1:
        mat.layer(i).blend = True
        mat.layer(i).blendsrc = GL_SRC_ALPHA
        mat.layer(i).blenddst = GL_ONE_MINUS_SRC_ALPHA
        mat.layer(i).depthWrite = GL_FALSE
        'mat.layer(i).alphatest = True
        'mat.layer(i).alpharef = 0.005
    Case 2:
        mat.layer(i).alphatest = True
        mat.layer(i).alpharef = 0.5
    End Select
End Sub

Private Sub SetAlpha(ByRef mat As bf2mat, ByVal i As Long)
    mat.layer(i).texcoff = 0
    mat.layer(i).texmapid = mat.texmapid(0)
    mat.layer(i).depthfunc = GL_LESS
    mat.layer(i).depthWrite = GL_TRUE
    mat.layer(i).blend = True
    mat.layer(i).blendsrc = GL_SRC_ALPHA
    mat.layer(i).blenddst = GL_ONE_MINUS_SRC_ALPHA
    mat.layer(i).lighting = False
End Sub

Private Sub SetAlphaTest(ByRef mat As bf2mat, ByVal i As Long)
    mat.layer(i).texcoff = 0
    mat.layer(i).texmapid = mat.texmapid(0)
    mat.layer(i).depthfunc = GL_LESS
    mat.layer(i).depthWrite = GL_TRUE
    mat.layer(i).alphatest = True
    mat.layer(i).alpharef = 0.5
    mat.layer(i).lighting = False
End Sub

Private Sub SetDetail(ByRef mat As bf2mat, ByVal i As Long)
    mat.layer(i).texcoff = 1
    mat.layer(i).texmapid = mat.texmapid(1)
    mat.layer(i).depthfunc = GL_EQUAL
    mat.layer(i).depthWrite = GL_FALSE
    mat.layer(i).blend = True
    mat.layer(i).blendsrc = GL_ZERO
    mat.layer(i).blenddst = GL_SRC_COLOR
    mat.layer(i).lighting = False
End Sub

Private Sub SetDirt(ByRef mat As bf2mat, ByVal i As Long)
    mat.layer(i).texcoff = 2
    mat.layer(i).texmapid = mat.texmapid(2)
    mat.layer(i).depthfunc = GL_EQUAL
    mat.layer(i).depthWrite = GL_FALSE
    mat.layer(i).blend = True
    mat.layer(i).blendsrc = GL_ZERO
    mat.layer(i).blenddst = GL_SRC_COLOR
    mat.layer(i).lighting = False
End Sub

Private Sub SetCrack(ByRef mat As bf2mat, ByVal i As Long)
    mat.layer(i).texcoff = 3
    mat.layer(i).texmapid = mat.texmapid(3)
    mat.layer(i).depthfunc = GL_EQUAL
    mat.layer(i).depthWrite = GL_FALSE
    mat.layer(i).blend = True
    mat.layer(i).blendsrc = GL_SRC_ALPHA
    mat.layer(i).blenddst = GL_ONE_MINUS_SRC_ALPHA
    mat.layer(i).lighting = True
End Sub


'swaps base (layer 1) and detail (layer 2) in case of alpha
Private Sub MakeAlpha(ByRef mat As bf2mat)
    With mat
        If .alphamode = 2 Then
            Dim tmp As Long
            tmp = .layer(1).texmapid
            .layer(1).texmapid = .layer(2).texmapid
            .layer(2).texmapid = tmp
            
            .layer(1).texcoff = 1
            .layer(2).texcoff = 0
            
            .layer(1).texmapid = .texmapid(1)
            .layer(2).texmapid = .texmapid(0)
            
            .layer(1).depthfunc = GL_LESS
            .layer(2).depthfunc = GL_EQUAL
            
            .layer(1).depthWrite = GL_TRUE
            .layer(2).depthWrite = GL_FALSE
            
            .layer(1).blend = False
            .layer(2).blend = True
            .layer(2).blendsrc = GL_ZERO
            .layer(2).blenddst = GL_SRC_COLOR
            
            .layer(1).alphatest = True
            .layer(1).alpharef = 0.5
        End If
    End With
End Sub


'builds shader
Public Sub BuildShader(ByRef mat As bf2mat, ByRef filename As String)
    With mat
        'reset
        .layernum = 0
        
        Select Case LCase(.fxfile)
        
        'SKINNEDMESH
        Case "skinnedmesh.fx"
            mat.glslprog = 0
            
            Select Case LCase(.technique)
            Case "alpha_test"
                .layernum = 1
                SetAlphaTest mat, 1
            Case Else
                .layernum = 1
                SetBase mat, .layernum
            End Select
            
        'BUNDLEDMESH
        Case "bundledmesh.fx"
            mat.glslprog = bundledmesh.prog
            mat.hasBump = False
            mat.hasWreck = False
            
            If .mapnum = 3 Then
                If InString(.map(1), "SpecularLUT") Then
                    mat.hasBump = False
                Else
                    mat.hasBump = True
                End If
            End If
            If .mapnum = 4 Then
                mat.hasBump = True
                mat.hasWreck = True
            End If
            
            'opaque
            .layernum = 1
            SetBase mat, 1
            
            'uvmap
            'If InStr(1, .technique, "AnimatedUV", vbTextCompare) > 0 Then
            '    .layer(1).texcoff = 1
            'End If
            
            'wreck (no bump)
            If .mapnum = 3 Then
                .layer(1).depthWrite = GL_TRUE
                
                .layernum = 2
                .layer(2).texcoff = 0
                .layer(2).texmapid = mat.texmapid(2)
                .layer(2).depthfunc = GL_EQUAL
                .layer(2).depthWrite = GL_FALSE
                If .alphamode = 1 Then .layer(2).depthfunc = GL_EQUAL 'note: does not render correctly, but we don't care
                If .alphamode = 2 Then .layer(2).depthfunc = GL_EQUAL
                .layer(2).blend = True
                .layer(2).blendsrc = GL_ZERO
                .layer(2).blenddst = GL_SRC_COLOR
                .layer(2).lighting = False
            End If
            
            'wreck
            If .mapnum = 4 Then
                .layer(1).depthWrite = GL_TRUE
                
                .layernum = 2
                .layer(2).texcoff = 0
                .layer(2).texmapid = mat.texmapid(3)
                .layer(2).depthfunc = GL_EQUAL
                .layer(2).depthWrite = GL_FALSE
                If .alphamode = 1 Then .layer(2).depthfunc = GL_EQUAL 'note: does not render correctly, but we don't care
                If .alphamode = 2 Then .layer(2).depthfunc = GL_EQUAL
                .layer(2).blend = True
                .layer(2).blendsrc = GL_ZERO
                .layer(2).blenddst = GL_SRC_COLOR
                .layer(2).lighting = False
            End If
            
        'STATICMESH
        Case "staticmesh.fx"
            mat.glslprog = 0
            
            'check if file is in vegetation directory
            Dim veggie As Boolean
            veggie = (InStr(1, filename, "vegitation", vbTextCompare) > 0)
            'todo: check texture file paths instead so file is displayed properly outside veggie dir?
            
            Select Case .technique
            
            'empty
            Case ""
                
            'misc
            Case "ColormapGloss", "EnvColormapGloss"
                .layernum = 1
                SetBase mat, 1
                
            Case "Alpha"
                .layernum = 1
                SetAlpha mat, 1
                
            Case "Alpha_Test"
                .layernum = 1
                SetAlphaTest mat, 1
                
            'statics
            Case "Base"
                If veggie Then
                    .layernum = 1
                    .layer(1).texcoff = 0
                    .layer(1).texmapid = mat.texmapid(0)
                    .layer(1).depthfunc = GL_LESS
                    .layer(1).depthWrite = GL_TRUE
                    .layer(1).alphatest = True
                    .layer(1).alpharef = 0.25
                    .layer(1).twosided = True
                Else
                    .layernum = 1
                    SetBase mat, 1
                End If
                
            Case "BaseDetail", _
                 "BaseDetailNDetail", _
                 "BaseDetailNDetailenvmap"
                
                If veggie Then
                    
                    .layernum = 2
                    
                    'detail (trunk texture)
                    mat.layer(1).texcoff = 1
                    mat.layer(1).texmapid = mat.texmapid(1)
                    mat.layer(1).depthfunc = GL_LESS
                    mat.layer(1).depthWrite = GL_TRUE
                    mat.layer(1).blend = False
                    mat.layer(1).lighting = True
                    
                    'base (trunk dirt)
                    mat.layer(2).texcoff = 0
                    mat.layer(2).texmapid = mat.texmapid(0)
                    mat.layer(2).depthfunc = GL_EQUAL
                    mat.layer(2).depthWrite = GL_FALSE
                    mat.layer(2).blend = True
                    mat.layer(2).blendsrc = GL_DST_COLOR
                    mat.layer(2).blenddst = GL_SRC_COLOR
                    mat.layer(2).lighting = False
                    
                Else
                    
                    .layernum = 2
                    SetBase mat, 1
                    SetDetail mat, 2
                    MakeAlpha mat
                    
                End If
                           
            Case "BaseDetailCrack", _
                 "BaseDetailCrackNCrack", _
                 "BaseDetailCrackNDetail", _
                 "BaseDetailCrackNDetailNCrack"
                
                .layernum = 3
                SetBase mat, 1
                SetDetail mat, 2
                SetCrack mat, 3
                
                .layer(1).texcoff = 0
                .layer(2).texcoff = 1
                'If vmesh.isbf2 Then
                    .layer(3).texcoff = 2
                'Else
                '    .layer(3).texcoff = 3
                'End If
                
                .layer(1).texmapid = .texmapid(0)
                .layer(2).texmapid = .texmapid(1)
                .layer(3).texmapid = .texmapid(2)
                
            Case "BaseDetailDirt", _
                 "BaseDetailDirtNDetail"
                 
                .layernum = 3
                SetBase mat, 1
                SetDetail mat, 2
                SetDirt mat, 3
                MakeAlpha mat
                
            Case "BaseDetailDirtCrack", _
                 "BaseDetailDirtCrackNCrack", _
                 "BaseDetailDirtCrackNDetail", _
                 "BaseDetailDirtCrackNDetailNCrack"
                 
                .layernum = 4
                SetBase mat, 1
                SetDetail mat, 2
                SetDirt mat, 4     'we swap dirt and crack for FH2
                SetCrack mat, 3    'we swap dirt and crack for FH2
                MakeAlpha mat
                
            'auto generate
            Case Else
                
                If InStr(1, .technique, "base", vbTextCompare) > 0 Then
                    .layernum = .layernum + 1
                    SetBase mat, .layernum
                    
                ElseIf InStr(1, .technique, "detail", vbTextCompare) > 0 Then
                    .layernum = .layernum + 1
                    SetDetail mat, .layernum
                    
                ElseIf InStr(1, .technique, "dirt", vbTextCompare) > 0 Then
                    .layernum = .layernum + 1
                    SetDirt mat, .layernum
                    
                ElseIf InStr(1, .technique, "crack", vbTextCompare) > 0 Then
                    .layernum = .layernum + 1
                    SetCrack mat, .layernum
                    
                ElseIf InStr(1, .technique, "humanskin", vbTextCompare) > 0 Then
                    .layernum = .layernum + 1
                    SetBase mat, .layernum
                
                Else
                    'all other cases (may be rendered incorrectly)
                    .layernum = 1
                    SetBase mat, .layernum
                End If
                
            End Select
            
            'texmap to UV offset lookup table
            Dim mapnum As Long
            Dim detail As Long
            Dim crack As Long
            If InStr(1, .technique, "Base", vbTextCompare) Then
                .mapuvid(mapnum) = mapnum
                mapnum = mapnum + 1
            End If
            If InStr(1, .technique, "Detail", vbTextCompare) Then
                .mapuvid(mapnum) = 1
                detail = mapnum
                mapnum = mapnum + 1
            End If
            If InStr(1, .technique, "Dirt", vbTextCompare) Then
                .mapuvid(mapnum) = 2
                mapnum = mapnum + 1
            End If
            If InStr(1, .technique, "Crack", vbTextCompare) Then
                .mapuvid(mapnum) = 3
                crack = mapnum
                mapnum = mapnum + 1
            End If
            If InStr(1, .technique, "NDetail", vbTextCompare) Then
                .mapuvid(mapnum) = detail
                mapnum = mapnum + 1
            End If
            If InStr(1, .technique, "NCrack", vbTextCompare) Then
                .mapuvid(mapnum) = crack
                mapnum = mapnum + 1
            End If
            
        End Select
    End With
End Sub
