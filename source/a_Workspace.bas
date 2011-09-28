Attribute VB_Name = "a_workspace"
Option Explicit

Private Type wspace_type
    loaded As Boolean
    filename As String
    
    fname_ske As String
    fname_mesh As String
    fname_anim As String
    fname_con As String
    
    'ske As bf2ske_file
    'mesh As bf2mesh
    'anim As baf_file
    
End Type
Public wspace  As wspace_type


'loads workspace from file
Public Function LoadWorkspace(ByVal filename As String)
    
    UnloadWorkspace
    
    'open file
    Dim ff As Integer
    ff = FreeFile
    Open filename For Input As #ff
    
    wspace.filename = filename
    
    Dim ln As String
    Dim str() As String
    Do Until EOF(ff)
        Line Input #ff, ln
        ln = Trim$(ln)
        If Len(ln) > 0 And Left$(ln, 1) <> ";" Then
            
            str() = Split(ln, "=", 2)
            
            Select Case str(0)
            Case "ske"
                wspace.fname_ske = str(1)
                LoadBF2Skeleton str(1)
            Case "mesh"
                wspace.fname_mesh = str(1)
                LoadBF2Mesh str(1)
            Case "anim"
                wspace.fname_anim = str(1)
                LoadBF2Anim str(1)
            Case "con"
                wspace.fname_con = str(1)
                LoadCon str(1)
            End Select
            
        End If
    Loop
    
    If opt_loadtextures Then LoadMeshTextures
    
    wspace.loaded = True
    
    'close file
    Close #ff
    
    'some hacks
    If vmesh.loadok Then
        If vmesh.isSkinnedMesh Then
            If vmesh.geomnum = 2 Then
                If InStr(LCase(wspace.fname_ske), "3p") Then
                    seldefault = MakeTag(1, 0, 0)
                End If
            End If
        End If
    End If
    
    LoadWorkspace = True
    Exit Function
errhandler:
    MsgBox "LoadWorkspace" & vbLf & err.description, vbCritical
End Function


'unloads workspace
Public Sub UnloadWorkspace()
    With wspace
        .loaded = False
        .filename = ""
    End With
    
    UnloadBF2Mesh
    UnloadMeshTextures
    UnloadBF2Anim
    UnloadBF2Skeleton
End Sub
