Attribute VB_Name = "a_Shader"
Option Explicit

'uniforms
Public nodetransform(40) As matrix4
Public nodetransformnum As Long


'shaders
Public bundledmesh As shader
Public skinnedmesh As shader


'loads all shaders
Public Sub LoadShaders()
    CreateProgram bundledmesh, xs("bundledmesh_vert.glsl"), xs("bundledmesh_frag.glsl")
    CreateProgram skinnedmesh, xs("skinnedmesh_vert.glsl"), xs("skinnedmesh_frag.glsl")
End Sub

'unloads all shaders
Public Sub UnloadShaders()
    DeleteProgram bundledmesh
    DeleteProgram skinnedmesh
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
Public Sub SetUniforms(ByRef sh As shader, ByRef mat As bf2mat)
Dim epw As float3
    epw.X = -eyeposworld.X 'DICE crap is flipped
    epw.y = eyeposworld.y
    epw.z = eyeposworld.z
    SetUniform3f sh, "eyeposworld", epw
    SetUniform1f sh, "hasBump", Bool2Float(mat.hasBump)
    SetUniform1f sh, "hasWreck", Bool2Float(mat.hasWreck)
    SetUniform1f sh, "hasXnimatedUV", Bool2Float(mat.hasAnimatedUV)
    SetUniform1f sh, "hasXlpha", Bool2Float(mat.alphamode > 0)
    SetUniform1f sh, "showLighting", Bool2Float(view_lighting)
    SetUniform1f sh, "showDiffuse", Bool2Float(view_textures)
    SetNodeTransforms sh, "nodetransform"
End Sub


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


'converts boolean to float
Public Function Bool2Float(ByVal v As Boolean) As Single
    If v Then Bool2Float = 1
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
