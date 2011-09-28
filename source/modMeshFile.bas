Attribute VB_Name = "a_MeshFile"
Option Explicit


'opens mesh file
Public Function OpenMeshFile(ByRef filename As String) As Boolean
    
    'check if file exists
    If Not FileExist(filename) Then
        MsgBox "File " & Chr(34) & " not found.", vbExclamation
        Exit Function
    End If
        
    SetStatus "info", "Loading " & filename & "..."
    Echo "File: " & Chr(34) & filename & Chr(34)
    Echo ""
    
    'determine type
    Dim r As Boolean
    Dim ext As String
    ext = LCase(GetFileExt(filename))
    Select Case ext
    Case "bundledmesh", "staticmesh", "skinnedmesh"
        r = LoadBF2Mesh(filename)
        If r Then
            If opt_loadtextures Then LoadMeshTextures
            If opt_loadsamples Then LoadSamplesFiles filename
            
            'auto-select geom1
            If vmesh.isSkinnedMesh Then
                If vmesh.geomnum = 2 Then
                    seldefault = MakeTag(1, 0, 0)
                End If
            End If
        End If
    Case "collisionmesh"
        r = LoadBF2Col(filename)
        If Not view_edges Then frmMain.mnuViewEdges_Click
    Case "samples", "samp_01", "samp_02", "samp_03"
        r = LoadSamples(bf2samples(0), filename)
    Case "sm"
        r = LoadStdMesh(filename)
        If r Then
            If opt_loadtextures Then
                LoadRS Replace(filename, ".sm", ".rs")
                StdMeshBindShaders
                LoadMeshTextures
            End If
        End If
    Case "tm"
        r = LoadTreeMesh(filename)
        If r Then
            If opt_loadtextures Then LoadMeshTextures
        End If
    Case "occ"
        r = LoadOccluder(filename)
    Case "obj"
        r = LoadOBJ(filename)
    Case "res"
        r = LoadFbMesh(filename)
    Case "geo"
        r = LoadFhxGeo(filename)
    Case "rig"
        r = LoadFhxRig(filename)
    Case "tri"
        r = LoadFhxTri(filename)
        If Not view_edges Then frmMain.mnuViewEdges_Click
    Case "ske"
        r = LoadBF2Skeleton(filename)
    Case "baf"
        r = LoadBF2Anim(filename)
    Case "bfmv"
        r = LoadWorkspace(filename)
    Case Else
        MsgBox "Unknown file type.", vbExclamation
        r = False
    End Select
    
    'success
    OpenMeshFile = r
End Function


'saves mesh file
Public Function SaveMeshFile(ByRef filename As String) As Boolean
    
Dim backup As Boolean
Dim backupname As String
    
    SetStatus "info", "Making backup of " & filename & "..."
    
    'make backup and delete existing file
    If FileExist(filename) Then
        backup = True
        backupname = filename & ".backup"
        FileCopy filename, backupname
        Kill filename
    End If
    
    SetStatus "info", "Saving " & filename & "..."
    
    'determine type
    Dim r As Boolean
    Dim ext As String
    ext = LCase(GetFileExt(filename))
    Select Case ext
    Case "bundledmesh", "staticmesh", "skinnedmesh"
        r = WriteVisMesh(filename)
    Case "sm"
        r = WriteStdMesh(filename)
    Case "samples", "samp_01", "samp_02", "samp_03"
        r = WriteSamples(bf2samples(0), filename)
    Case Else
        MsgBox "Unknown file type.", vbExclamation
        r = False
    End Select
    
    'delete/restore backup
    If backup Then
        If r Then
            'delete backup
            Kill backupname
        Else
            'restore backup
            Kill filename
            FileCopy backupname, filename
        End If
    End If
    
    'success
    SaveMeshFile = r
End Function


'unloads mesh file
Public Sub CloseMeshFile()
    
    UnloadWorkspace
    
    'bf2
    UnloadBF2Mesh
    UnloadBF2Skeleton
    UnloadBF2Col
    UnloadBF2Anim
    UnloadSamplesFiles
    UnloadSamples bf2samples(0)
    UnloadOccluder
    UnloadCon
    
    'bf42
    UnloadStdMesh
    UnloadRS
    UnloadTreeMesh
    
    'misc
    UnloadObj
    
    'FrostBite
    UnloadFbMesh
    
    'fhx
    UnloadFhxGeo
    UnloadFhxTri
    UnloadFhxRig
    
    'misc
    UnloadMeshTextures
End Sub
