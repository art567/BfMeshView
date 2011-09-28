Attribute VB_Name = "BF1942_Shader"
Option Explicit

Private inblock As Boolean


Private Type subshader_type
    name As String
    ustring As String 'e.g. "StandardMesh/Default", not sure what it's for
    lighting As Boolean
    lightingSpecular As Boolean
    envmap As Boolean
    transparent As Boolean
    twosided As Boolean
    depthWrite As Boolean
    alphaTestRef As Single
    materialDiffuse As color3f
    materialSpecular As color3f
    materialSpecularPower As Single
    texture As String
    
    'internal
    texmapid As Long
End Type

Private Type stdshader_type
    filename As String
    loaded As Boolean
    
    subshader() As subshader_type 'range:[1 to subshader_num]
    subshader_num As Long
End Type

Public stdshader As stdshader_type


'binds shaders on standard mesh
Public Sub StdMeshBindShaders()
    Dim s As Long
    Dim i As Long
    Dim j As Long
    For s = 1 To stdshader.subshader_num
        With stdmesh
            For i = 0 To .lodnum - 1
                With .lod(i)
                    For j = 0 To .matnum - 1
                        With .mat(j)
                            If LCase(.matname) = LCase(stdshader.subshader(s).name) Then
                                .shaderid = s
                            End If
                        End With
                    Next j
                End With
            Next i
        End With
    Next s
End Sub


Public Sub BindStdShader(ByVal id As Long)
    With stdshader.subshader(id)
        'note: view_textures is true within the scope of this function
        
        If .texmapid > 0 Then
            glColor3fv .materialDiffuse.r
            BindTexture .texmapid
        Else
            glColor3f 1, 0.25, 0.25
            UnbindTexture
        End If
        
        If .lighting And view_lighting Then
            glEnable GL_LIGHTING
        Else
            glDisable GL_LIGHTING
        End If
        
        If .twosided Or view_backfaces Then
            glDisable GL_CULL_FACE
        Else
            glEnable GL_CULL_FACE
        End If
        
        If .transparent Then
            glBlendFunc GL_SRC_ALPHA, GL_ONE_MINUS_SRC_ALPHA
            glEnable GL_BLEND
        Else
            glDisable GL_BLEND
        End If
        
        If .depthWrite Then
            glDepthMask True
        Else
            glDepthMask False
        End If
        
        If .alphaTestRef > 0 Then
            glEnable GL_ALPHA_TEST
            glAlphaFunc GL_GREATER, .alphaTestRef
        Else
            glDisable GL_ALPHA_TEST
        End If
        
    End With
End Sub


'processes line
Private Sub ParseLine(ByRef ln As String)
    
    'remove evil tabs
    ln = Replace(ln, Chr(9), " ")
    
    'remove whitespaces
    ln = Trim$(ln)
    
    'remove excessive spaces
    ln = Replace(ln, "  ", " ")
    
    'remove ";"
    ln = Replace(ln, ";", "")
    
    'skip comments
    If Len(ln) = 0 Then Exit Sub
    
    'process
    Dim str() As String
    str() = Split(ln, " ")
    
    If Not inblock Then
        Select Case LCase(str(0))
        Case "subshader"
            
            With stdshader
                .subshader_num = .subshader_num + 1
                ReDim Preserve .subshader(1 To .subshader_num)
            End With
            
            With stdshader.subshader(stdshader.subshader_num)
                .name = Replace(str(1), Chr(34), "")
                .ustring = Replace(str(2), Chr(34), "")
                
                'default settings
                .materialDiffuse = color3f(1, 1, 1)
                .materialSpecular = color3f(1, 1, 1)
                .materialSpecularPower = 15
                .lighting = True
                .lightingSpecular = False
                .transparent = False
                .twosided = False
                .envmap = False
                .depthWrite = True
                .alphaTestRef = 0
                
                .texmapid = 0
            End With
            
        Case "{"
            inblock = True
            
        Case Else
            Echo "Unknown command: " & str(0)
        End Select
    Else
        With stdshader.subshader(stdshader.subshader_num)
            Select Case LCase(str(0))
            Case "lighting":              .lighting = str(1)
            Case "lightingspecular":      .lightingSpecular = str(1)
            Case "twosided":              .twosided = str(1)
            Case "envmap":                .envmap = str(1)
            Case "depthwrite":            .depthWrite = str(1)
            Case "transparent":           .transparent = str(1)
            Case "alphatestref":          .alphaTestRef = val(str(1))
            Case "texture":               .texture = Replace(str(1), Chr(34), "")
            Case "materialspecularpower": .materialSpecularPower = val(str(1))
            Case "materialdiffuse"
                .materialDiffuse.r = val(str(1))
                .materialDiffuse.g = val(str(2))
                .materialDiffuse.b = val(str(3))
            Case "materialspecular"
                .materialSpecular.r = val(str(1))
                .materialSpecular.g = val(str(2))
                .materialSpecular.b = val(str(3))
            Case "}"
                inblock = False
            Case Else
                Echo "Unknown property: " & str(0)
            End Select
        End With
    End If
    
End Sub


'loads BF1942 shader file
Public Sub LoadRS(ByRef filename As String)
    On Error GoTo errorhandler
    
    'check if file exists
    If Not FileExist(filename) Then
        Echo "File " & Chr(34) & filename & Chr(34) & " not found."
        Exit Sub
    End If
    
    Echo "--------------------------------------------------------------------"
    Echo "Loading " & filename
    Echo ""
    
    'open file
    Dim ff As Integer
    ff = FreeFile
    Open filename For Input As #ff
    
    stdshader.filename = filename
    stdshader.loaded = True
    
    'parse
    Dim ln As String
    Dim linenum As Long
    Do Until EOF(ff)
        linenum = linenum + 1
        Line Input #ff, ln
        
        Dim lnarr() As String
        lnarr = Split(ln, vbLf)
        
        Dim i As Long
        For i = LBound(lnarr()) To UBound(lnarr())
            ParseLine lnarr(i)
        Next i
        
    Loop
    
    'close file
    Close ff
    
    'success
    Exit Sub
errorhandler:
    Echo "Shader parse error on line " & linenum & ":" & err.Description
    Echo ">>> " & ln
End Sub


'unloads stuff
Public Sub UnloadRS()
    With stdshader
        .loaded = False
        .filename = ""
        
        .subshader_num = 0
        Erase .subshader()
    End With
End Sub


'fills treeview hierarchy
Public Sub FillTreeStdMeshShaders(ByRef tree As MSComctlLib.TreeView)
    Dim i As Long
    Dim j As Long
    With stdshader
        If Not .loaded Then Exit Sub
        
        Dim n As MSComctlLib.node
        
        Dim tag As Long
        tag = MakeTag(2, 0, 0)
        
        Dim rootname As String
        rootname = "stdshader_root"
        Set n = tree.Nodes.Add(, tvwChild, rootname, GetFileName(.filename), "file")
        n.Expanded = True
        n.tag = tag
        
        'lods
        For i = 1 To .subshader_num
            With .subshader(i)
                
                Dim matname As String
                matname = "stdshadermat" & i
                Set n = tree.Nodes.Add(rootname, tvwChild, matname, .name, "mat")
                n.tag = tag
                
                Set n = tree.Nodes.Add(matname, tvwChild, matname & "|lighting", "Lighting: " & .lighting, "prop")
                n.tag = tag
                
                Set n = tree.Nodes.Add(matname, tvwChild, matname & "|spec", "Specular: " & .lightingSpecular, "prop")
                n.tag = tag
                
                Set n = tree.Nodes.Add(matname, tvwChild, matname & "|twosided", "TwoSided: " & .twosided, "prop")
                n.tag = tag
                
                Set n = tree.Nodes.Add(matname, tvwChild, matname & "|envmap", "EnvMap: " & .envmap, "prop")
                n.tag = tag
                
                Set n = tree.Nodes.Add(matname, tvwChild, matname & "|transparent", "Transparent: " & .transparent, "prop")
                n.tag = tag
                
                Set n = tree.Nodes.Add(matname, tvwChild, matname & "|alphatest", "AlphaTest: " & .alphaTestRef, "prop")
                n.tag = tag
                
                Set n = tree.Nodes.Add(matname, tvwChild, matname & "|depthwrite", "DepthWrite: " & .depthWrite, "prop")
                n.tag = tag
                
                Dim texname As String
                Dim icon As String
                If .texmapid Then
                    icon = "tex"
                    texname = GetFileName(texmap(.texmapid).filename)
                Else
                    texname = .texture
                    icon = "texmissing"
                End If
                Set n = tree.Nodes.Add(matname, tvwChild, matname & "|tex", texname, icon)
                n.tag = MakeTag(2, i, 0)
                
            End With
        Next i
        
    End With
End Sub

