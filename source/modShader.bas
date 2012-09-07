Attribute VB_Name = "a_Shader"
Option Explicit

'uniforms
Public nodetransform(40) As matrix4
Public nodetransformnum As Long


'shaders
Public bundledmesh As shader
Public skinnedmesh As shader
Public staticmesh As shader
Public trunk As shader
Public leaf As shader


'loads all shaders
Public Sub LoadShaders()
    CreateProgram bundledmesh, xs("bundledmesh_vert.glsl"), xs("bundledmesh_frag.glsl")
    CreateProgram skinnedmesh, xs("skinnedmesh_vert.glsl"), xs("skinnedmesh_frag.glsl")
    CreateProgram staticmesh, xs("staticmesh_vert.glsl"), xs("staticmesh_frag.glsl")
    CreateProgram trunk, xs("trunk_vert.glsl"), xs("trunk_frag.glsl")
    CreateProgram leaf, xs("leaf_vert.glsl"), xs("leaf_frag.glsl")
End Sub

'unloads all shaders
Public Sub UnloadShaders()
    DeleteProgram bundledmesh
    DeleteProgram skinnedmesh
    DeleteProgram staticmesh
    DeleteProgram trunk
    DeleteProgram leaf
End Sub


'reloads all shader
Public Sub ReloadShaders()
    UnloadShaders
    LoadShaders
    BuildFFPShaders
End Sub


'tiny helper
Private Function xs(ByRef fname As String) As String
    xs = LoadTextFile_NoError(App.path & "\shaders\" & fname)
End Function

'...
Public Sub SetUniforms(ByRef prog As GLuint, ByRef mat As bf2mat)
    
    'texture handles
    Dim i As Long
    For i = 0 To 7
        SetUniform1i prog, "texture" & i, i
    Next i
    SetUniform1i prog, "envmap", envmapChan
    
    'uniforms
    SetUniform3f prog, "eyeposworld", FlipX(eyeposworld)
    SetUniform3f prog, "eyevecworld", FlipX(eyevecworld)
    SetUniform1i prog, "hasBump", Bool2Int(mat.hasBump And view_bumpmap)
    SetUniform1i prog, "hasWreck", Bool2Int(mat.hasWreck)
    SetUniform1i prog, "hasAnimatedUV", Bool2Int(mat.hasAnimatedUV)
    SetUniform1i prog, "hasAlpha", Bool2Int(mat.alphamode > 0)
    SetUniform1i prog, "hasBumpAlpha", Bool2Int(mat.hasBumpAlpha)
    SetUniform1i prog, "hasDirt", Bool2Int(mat.hasDirt)
    SetUniform1i prog, "hasCrack", Bool2Int(mat.hasCrack)
    SetUniform1i prog, "hasCrackN", Bool2Int(mat.hasCrackN And view_bumpmap)
    SetUniform1i prog, "hasDetailN", Bool2Int(mat.hasDetailN And view_bumpmap)
    SetUniform1i prog, "hasEnvMap", Bool2Int(mat.hasEnvMap And view_envmap)
    SetUniform1i prog, "showLighting", Bool2Int(view_lighting)
    SetUniform1i prog, "showDiffuse", Bool2Int(view_textures)
    SetUniform1f prog, "timephase", timephase
    SetNodeTransforms prog, "nodetransform"
End Sub

Private Sub AddReportVar(ByRef str As String, ByRef name As String)
    Dim v As Long
    v = 666
    Dim loc As Long
    loc = glGetUniformLocation(bundledmesh.prog, name)
    glGetUniformiv bundledmesh.prog, loc, v
    str = str & name & " @ " & loc & ": " & v & vbLf
End Sub

Public Sub GetReport()
    Dim str As String
    AddReportVar str, "texture0"
    AddReportVar str, "texture1"
    AddReportVar str, "texture2"
    AddReportVar str, "texture3"
    AddReportVar str, "texture4"
    AddReportVar str, "texture5"
    AddReportVar str, "texture6"
    AddReportVar str, "texture7"
    AddReportVar str, "hasBump"
    AddReportVar str, "hasWreck"
    AddReportVar str, "hasAnimatedUV"
    AddReportVar str, "hasAlpha"
    AddReportVar str, "showLighting"
    AddReportVar str, "showDiffuse"
    MsgBox str
End Sub

'DICE crap is flipped on X axis
Public Function FlipX(ByRef v As float3) As float3
    FlipX = float3(-v.x, v.y, v.z)
End Function

'converts boolean to float
Public Function Bool2Float(ByVal v As Boolean) As Single
    If v Then Bool2Float = 1
End Function

'converts boolean to int
Public Function Bool2Int(ByVal v As Boolean) As Long
    If v = True Then
        Bool2Int = 1
    Else
        Bool2Int = 0
    End If
End Function


'loads entire file as text
Private Function LoadTextFile_NoError(ByRef fname As String) As String
On Error GoTo errhandler
    Dim str As String
    Dim ff As Integer
    ff = FreeFile
    Open fname For Input As ff
    LoadTextFile_NoError = StrConv(InputB(LOF(ff), ff), vbUnicode)
    Close ff
    Exit Function
errhandler:
    'MsgBox "LoadTextFile_NoError" & vbLf & err.description, vbCritical
End Function


'rebuilds shader stuff
Public Sub BuildFFPShaders()
    With vmesh
        If Not .loadok Then Exit Sub
        Dim i As Long
        For i = 0 To .geomnum - 1
            With .geom(i)
                Dim j As Long
                For j = 0 To .lodnum - 1
                    With .lod(j)
                        Dim k  As Long
                        For k = 0 To .matnum - 1
                            BuildShader .mat(k), vmesh.filename
                        Next k
                    End With
                Next j
            End With
        Next i
    End With
End Sub
