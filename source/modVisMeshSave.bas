Attribute VB_Name = "BF2_MeshSave"
Option Explicit


'writes mesh to file
Public Function WriteVisMesh(ByRef filename As String) As Boolean
    On Error GoTo errorhandler
    
Dim i As Long
Dim j As Long
    
    'open file
    Dim ff As Integer
    ff = FreeFile
    Open filename For Binary As #ff
    
    With vmesh
        
        'header (20 bytes)
        Put #ff, , .head
        
        'unknown (1 byte)
        Put #ff, , .u1
        
        '--- geom table ---------------------------------------------------------------------------
        
        'geomnum (4 bytes)
        Put #ff, , .geomnum
        
        'geoms
        For i = 0 To .geomnum - 1
            
            'meshnum (4 bytes)
            Put #ff, , .geom(i).lodnum
            
        Next i
        
        '--- vertex attribute table ---------------------------------------------------------------
        
        'snum (4 bytes)
        Put #ff, , .vertattribnum
        
        'sdata (4bytes * snum)
        Put #ff, , .vertattrib()
        
        '--- vertices -----------------------------------------------------------------------------
        
        Put #ff, , .vertformat  '4 bytes
        Put #ff, , .vertstride  '4 bytes
        Put #ff, , .vertnum     '4 bytes
        Put #ff, , .vert()      '? bytes
        
        '--- indices ------------------------------------------------------------------------------
        
        Put #ff, , .indexnum    '4 bytes
        Put #ff, , .Index()     '? bytes
        
        '--- transforms ---------------------------------------------------------------------------
        
        If Not vmesh.isSkinnedMesh Then
            
            'unknown (4 bytes)
            Put #ff, , .u2
            
        End If
        
        '? bytes
        For i = 0 To .geomnum - 1
            For j = 0 To .geom(i).lodnum - 1
                WriteBF2MeshLodRig ff, .geom(i).lod(j)
            Next j
        Next i
        
        '--- groups ------------------------------------------------------------------------------
        
        For i = 0 To .geomnum - 1
            For j = 0 To .geom(i).lodnum - 1
                WriteBF2MeshLod ff, .geom(i).lod(j)
            Next j
        Next i
        
        '--- end of file -------------------------------------------------------------------------
        
    End With
    
    'close file
    Close #ff
    
    WriteVisMesh = True
    Exit Function
errorhandler:
    MsgBox "WriteVisMesh" & vbLf & err.description, vbCritical
End Function


'writes mesh bounds
Private Sub WriteBF2MeshLodRig(ByRef ff As Integer, ByRef lod As bf2_lod)
Dim i As Long
Dim j As Long
    With lod
        
        'bounds (24 bytes)
        Put #ff, , .min
        Put #ff, , .max
        
        'unknown (12 bytes)
        If vmesh.head.version <= 6 Then
            Put #ff, , .pivot
        End If
        
        If vmesh.isSkinnedMesh Then
            
            'rignum (4 bytes)
            Put #ff, , .rignum
            
            'rig data (? bytes)
            For i = 0 To .rignum - 1
                
                'bone num (4 bytes)
                Put #ff, , .rig(i).bonenum
                
                'bone data (68 bytes * bonenum)
                For j = 0 To .rig(i).bonenum - 1
                    Put #ff, , .rig(i).bone(j).id
                    Put #ff, , .rig(i).bone(j).matrix
                Next j
                
            Next i
            
        Else
            
            'nodenum (4 bytes)
            Put #ff, , .nodenum
            
            'node data (64 bytes * nodenum)
            If Not vmesh.isBundledMesh Then
                
                Put #ff, , .node()
                
            End If
            
        End If
        
    End With
End Sub


'writes string
Private Sub WriteBF2String(ByRef ff As Integer, ByRef str As String)
Dim strlen As Long
    strlen = Len(str)
    
    'write length (4 bytes)
    Put #ff, , strlen
    
    'write characters
    Dim i As Long
    For i = 1 To strlen
        Dim b As Byte
        b = Asc(Mid(str, i, 1))
        Put #ff, , b
    Next i
End Sub


'writes material chunk
Private Sub WriteBF2MeshMat(ByRef ff As Integer, ByRef mat As bf2_mat)
Dim i As Long
    With mat
        
        'alphamode (4 bytes)
        If Not vmesh.isSkinnedMesh Then
            Put #ff, , .alphamode
        End If
        
        'fx filename
        WriteBF2String ff, .fxfile
        
        'material name
        WriteBF2String ff, .technique
        
        'mapnum (4 bytes)
        Put #ff, , .mapnum
        
        'map[]
        For i = 0 To .mapnum - 1
            WriteBF2String ff, .map(i)
        Next i
        
        'geometry info
        Put #ff, , .vstart    '4 bytes
        Put #ff, , .istart    '4 bytes
        Put #ff, , .inum      '4 bytes
        Put #ff, , .vnum      '4 bytes
        
        'unknown
        Put #ff, , .u4      '4 bytes
        Put #ff, , .u5      '4 bytes
        
        'bounds
        If Not vmesh.isSkinnedMesh Then
            If vmesh.head.version = 11 Then
                Put #ff, , .mmin
                Put #ff, , .mmax
            End If
        End If
                
    End With
End Sub


'writes LOD chunk
Private Sub WriteBF2MeshLod(ByRef ff As Integer, ByRef lod As bf2_lod)
Dim i As Long
    With lod
        
        'matnum (4 bytes)
        Put #ff, , .matnum
        
        '? bytes
        For i = 0 To .matnum - 1
            WriteBF2MeshMat ff, .mat(i)
        Next i
        
    End With
End Sub
