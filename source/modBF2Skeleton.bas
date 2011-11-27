Attribute VB_Name = "BF2_Skeleton"
Option Explicit

Public fooz As Long '???

Private Type skenode
    name As String
    parent As Integer
    rot As quat
    pos As float3
    
    'internal
    localmat As matrix4
    localmatanim As matrix4
    worldmat As matrix4
    worldmatbackup As matrix4
End Type
Public Type bf2ske_file
    version As Long
    
    nodenum As Long
    node() As skenode
    
    'internal
    cambone As Long
    filename As String
    loaded As Boolean
End Type
Public bf2ske As bf2ske_file


'reads string from file
Private Function ReadBF2SkeletonString(ByRef ff As Integer)
    
    'read number of characters (2 bytes)
    Dim num As Integer
    Get #ff, , num
    If num = 0 Then Exit Function
    
    'read characters (num)
    'note: this includes a string-terminator
    Dim chars() As Byte
    ReDim chars(0 To num - 1)
    Get #ff, , chars()
    
    'return save string
    ReadBF2SkeletonString = SafeString(chars, num - 1)
End Function


'loads skeleton from file
Public Function LoadBF2Skeleton(ByVal filename As String) As Boolean
    On Error GoTo errorhandler
    
    If Not FileExist(filename) Then
        MsgBox "File " & filename & " not found.", vbExclamation
        Exit Function
    End If
    
    Dim i As Long
    
    'open file
    Dim ff As Integer
    ff = FreeFile
    Open filename For Binary Access Read Lock Write As #ff
    
    With bf2ske
        .loaded = False
        .filename = filename
        .cambone = -1
        
        'read version (4 bytes)
        Get #ff, , .version
        Echo "version: " & .version
        
        'read nodenum (4 bytes)
        Get #ff, , .nodenum
        Echo "nodenum: " & .nodenum
        
        'read nodes
        ReDim .node(0 To .nodenum - 1)
        For i = 0 To .nodenum - 1
            With .node(i)
                
                .name = ReadBF2SkeletonString(ff)
                
                Echo "node[" & i & "]: " & .name
                
                Get #ff, , .parent
                Get #ff, , .rot
                Get #ff, , .pos
                               
                mat4identity .localmat
                mat4identity .worldmat
                
                'correct rotation
                'bf2 stores the inverse of what sanity tells us
                .rot.X = -.rot.X
                .rot.Y = -.rot.Y
                .rot.z = -.rot.z
                
                'correct position
                'mesh nodes have garbage coordinates
                If Magnitude(.pos) > 999 Then
                    .pos = float3(0, 0, 0)
                    QuatIdentity .rot
                End If
                
                'QuatNormalize .rot
                
                mat4setrot .localmat, .rot
                mat4setpos .localmat, .pos
                
            End With
        Next i
        
        For i = 0 To .nodenum - 1
            With .node(i)
                Echo "rot: " & .rot.X & "," & .rot.Y & "," & .rot.z & "," & .rot.w
            End With
        Next i
        
        For i = 0 To .nodenum - 1
            With .node(i)
                Echo "pos: " & .pos.X & "," & .pos.Y & "," & .pos.z
            End With
        Next i

        '--- end of file ------------------------------------------------------------------
        
        Echo "done reading " & loc(ff)
        Echo "file size is " & LOF(ff)
        Echo ""
        
        'compute world space bone matrices world space
        For i = 0 To .nodenum - 1
            Dim p As Long
            p = .node(i).parent
            
            If p = -1 Then
                'root
                .node(i).worldmat = .node(i).localmat
            Else
                'child
                .node(i).worldmat = mat4mult(.node(i).localmat, .node(p).worldmat)
            End If
            
            'reset default animation transform and make world space backup
            .node(i).localmatanim = .node(i).localmat
            .node(i).worldmatbackup = .node(i).worldmat
            
            'detect camera bone
            If .node(i).name = "Camerabone" Then
                .cambone = i
            End If
        Next i
        
        .loaded = True
    End With
    
    'close file
    Close #ff
        
    'success
    LoadBF2Skeleton = True
    Exit Function
errorhandler:
    MsgBox "LoadBF2Skeleton" & vbLf & err.description, vbCritical
End Function


'reset skeleton pose
Public Sub ResetBF2Skeleton()
    Dim i As Long
    With bf2ske
        If Not .loaded Then Exit Sub
        
        For i = 0 To .nodenum - 1
            .node(i).localmatanim = .node(i).localmat
            .node(i).worldmat = .node(i).worldmatbackup
        Next i
    End With
End Sub


'deform skeleton
Public Sub DeformBF2Skeleton(ByRef anim As baf_file, ByVal frame As Long)
    Dim i As Long
    Dim boneIndex As Long
    Dim pos As float3
    Dim rot As quat
    With bf2ske
        If Not .loaded Then Exit Sub
        If Not anim.loaded Then Exit Sub
        
        'clamp frame to animation range
        If frame < 0 Then frame = 0
        If frame > anim.framenum - 1 Then frame = anim.framenum - 1
        
        'reset world matrices so that un-animated bones are always in default pose
        ResetBF2Skeleton
        
        'fill local animation transform for each animated bone
        For i = 0 To anim.bonenum - 1
            boneIndex = anim.boneId(i)
            
            If boneIndex < .nodenum Then
                pos = anim.boneData(i).frame(frame).pos
                rot = anim.boneData(i).frame(frame).rot
                
                mat4identity .node(boneIndex).localmatanim
                mat4setpos .node(boneIndex).localmatanim, pos
                mat4setrot .node(boneIndex).localmatanim, rot
            End If
        Next i
        
        'transform to world space
        For i = 0 To .nodenum - 1
            Dim p As Long
            p = .node(i).parent
            
            If p = -1 Then
                'root
                .node(i).worldmat = .node(i).localmatanim
            Else
                'child
                .node(i).worldmat = mat4mult(.node(i).localmatanim, .node(p).worldmat)
            End If
        Next i
    End With
End Sub


'fill treeview
Public Sub FillTreeBF2Skeleton(ByRef tree As MSComctlLib.TreeView)
    With bf2ske
        On Error GoTo errhandler
        If Not .loaded Then Exit Sub
        
        Dim n As MSComctlLib.node
        
        'file root
        Dim rootname As String
        rootname = "ske_root"
        Set n = tree.Nodes.Add(, , rootname, GetFileName(.filename), "file")
        n.Expanded = True
        n.tag = 0
        
        'version leaf
        Set n = tree.Nodes.Add(rootname, tvwChild, rootname & "|ver", "Version: " & .version, "prop")
        n.tag = 0
        
        'nodenum leaf
        Set n = tree.Nodes.Add(rootname, tvwChild, rootname & "|nodenum", "Nodes: " & .nodenum, "prop")
        n.tag = 0
        
        'loop nodes
        Dim i As Long
        For i = 0 To .nodenum - 1
            With .node(i)
                
                If .parent = -1 Then
                    'root
                    Set n = tree.Nodes.Add(rootname, tvwChild, "node" & i, .name, "lod")
                    n.Expanded = False
                Else
                    'child
                    Set n = tree.Nodes.Add("node" & .parent, tvwChild, "node" & i, .name, "lod")
                    n.Expanded = True
                End If
                n.tag = i
                
            End With
        Next i
        
    End With
    Exit Sub
errhandler:
    MsgBox "FillTreeColMesh" & vbLf & err.description, vbCritical
End Sub


'draws skeleton
Public Sub DrawBF2Skeleton()
Dim i As Long
Dim p As Long
    With bf2ske
        If Not .loaded Then Exit Sub
        
        If vmesh.loadok Then
            If Not view_bonesys Then Exit Sub
        End If
        
        glEnable GL_CULL_FACE
        glDisable GL_TEXTURE_2D
        glDisable GL_BLEND
        glDisable GL_ALPHA_TEST
        
        glShadeModel GL_FLAT
        glEnable GL_LIGHTING
        glEnable GL_DEPTH_TEST
        For i = 0 To .nodenum - 1
            p = .node(i).parent
            
            'draw bone
            If p > -1 Then
                glPushMatrix
                    'glMultMatrixf .node(p).worldmat.m(0)
                    
                    Dim foo As matrix4
                    mat4identity foo
                    mat4lookat foo, SubFloat3(mat4getpos(.node(p).worldmat), mat4getpos(.node(i).worldmat)), float3(0, 1, 0)
                    mat4setpos foo, mat4getpos(.node(p).worldmat)
                    glMultMatrixf foo.m(0)
                    
                    Dim Length As Single
                    Length = Distance(mat4getpos(.node(p).worldmat), mat4getpos(.node(i).worldmat))
                    
                    glColor3f 0.75, 0.75, 0.75
                    DrawBone2 Length / 8, Length
                glPopMatrix
            End If
        Next i
        glDisable GL_DEPTH_TEST
        glDisable GL_LIGHTING
        glShadeModel GL_SMOOTH
        
        StartAAPoint 5
        StartAALine 1.3
        
        For i = 0 To .nodenum - 1
            p = .node(i).parent
            
            glColor3f 1, 1, 0
            
            'draw dot
            glBegin GL_POINTS
                glVertex3fv .node(i).worldmat.m(12)
            glEnd
            
            'draw line
            If p > -1 Then
                glBegin GL_LINES
                    glVertex3fv .node(p).worldmat.m(12)
                    glVertex3fv .node(i).worldmat.m(12)
                glEnd
            End If
            
            'draw pivot
            glPushMatrix
                glMultMatrixf .node(i).worldmat.m(0)
                
                DrawPivot 0.01
            glPopMatrix
            
        Next i
        EndAALine
        EndAAPoint
        
    End With
End Sub


Private Sub DrawBone(ByVal w As Single, ByVal L As Single)
    If L < w * 2 Then L = w * 2
    
    glBegin GL_TRIANGLE_FAN
        glVertex3f 0, 0, 0
        glVertex3f 0, w, -w
        glVertex3f w, w, 0
        glVertex3f 0, w, w
        glVertex3f -w, w, 0
        glVertex3f 0, w, -w
    glEnd
    glBegin GL_TRIANGLE_FAN
        glVertex3f 0, L, 0
        glVertex3f -w, w, 0
        glVertex3f 0, w, w
        glVertex3f w, w, 0
        glVertex3f 0, w, -w
        glVertex3f -w, w, 0
    glEnd
End Sub

Private Sub DrawBone2(ByVal w As Single, ByVal L As Single)
    If L < w * 2 Then L = w * 2
    
    glEnable GL_NORMALIZE
    glBegin GL_TRIANGLE_FAN
        glVertex3f 0, 0, 0
        glVertex3f 0, -w, w
        
        glNormal3f -1, -1, -1
        glVertex3f -w, 0, w
        
        glNormal3f -1, 1, -1
        glVertex3f 0, w, w
        
        glNormal3f 1, 1, -1
        glVertex3f w, 0, w
        
        glNormal3f 1, -1, -1
        glVertex3f 0, -w, w
    glEnd
    glBegin GL_TRIANGLE_FAN
        glVertex3f 0, 0, L
        glVertex3f w, 0, w
        
        glNormal3f 1, 1, 1
        glVertex3f 0, w, w
        
        glNormal3f -1, 1, 1
        glVertex3f -w, 0, w
        
        glNormal3f -1, -1, 1
        glVertex3f 0, -w, w
        
        glNormal3f 1, -1, 1
        glVertex3f w, 0, w
    glEnd
    glDisable GL_NORMALIZE
End Sub


Public Sub UnloadBF2Skeleton()
    With bf2ske
        .loaded = False
        .filename = ""
        
        .nodenum = 0
        Erase .node()
    End With
End Sub
