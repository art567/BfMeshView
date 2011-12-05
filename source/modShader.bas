Attribute VB_Name = "a_Shader"
Option Explicit

'uniforms
Public nodetransform(40) As matrix4
Public nodetransformnum As Long


'shaders
Public bundledmesh As shader


'loads all shaders
Public Sub LoadShaders()
    
    CreateProgram bundledmesh, xs("bundledmesh_vert.glsl"), xs("bundledmesh_frag.glsl")
    
End Sub


'reloads all shader
Public Sub ReloadShaders()
    UnloadShaders
    LoadShaders
    BuildFFPShaders
End Sub


'unloads all shaders
Public Sub UnloadShaders()
    DeleteProgram bundledmesh
End Sub


'tiny helper
Private Function xs(ByRef fname As String) As String
    xs = LoadTextFile_NoError(App.path & "\shaders\" & fname)
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


'converts boolean to float
Public Function Bool2Float(ByRef v As Boolean) As Single
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
