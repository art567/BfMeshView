Attribute VB_Name = "FHX_Rig"
Option Explicit

Private Type fhxbone
    nodeid As Integer
    parent As Integer
    pos As float3
    rot As quat
    
    'internal
    mat As matrix4
End Type

Private Type fhxrig_file
    head As fileheader
    
    nodenum As Long
    node() As fhxbone
    
    'internal
    filename As String
    loaded As Boolean
End Type

Public fhxrig As fhxrig_file


'loads rig from file
Public Function LoadFhxRig(ByVal filename As String) As Boolean
    On Error GoTo errorhandler
    
    'open file
    Dim ff As Integer
    ff = FreeFile
    Open filename For Binary Access Read Lock Write As #ff
    
    Dim i As Long
    With fhxrig
        .loaded = False
        .filename = filename
        
        'read header (16 bytes)
        Get #ff, , .head
        Echo "file format version: " & .head.version
        Echo "size: " & .head.size & "/" & LOF(ff)
        Echo "offset: " & .head.offset
        Echo ""
        
        If .head.version < 2 Then
            MsgBox "File version no longer supported.", vbExclamation
            Close ff
            Exit Function
        End If
        
        'read nodenum (4 bytes)
        Get #ff, , .nodenum
        Echo "nodenum: " & .nodenum
        
        'read nodes
        ReDim .node(0 To .nodenum - 1)
        For i = 0 To .nodenum - 1
            With .node(i)
                
                Get #ff, , .nodeid
                Get #ff, , .parent
                Get #ff, , .pos
                Get #ff, , .rot
                
                Echo "node " & i & ": id:" & .nodeid & " parent: " & .parent
                
                mat4identity .mat
                mat4setpos .mat, .pos
                mat4setrot .mat, .rot
                If .parent > -1 Then
                    .mat = mat4mult(.mat, fhxrig.node(.parent).mat)
                End If
                
            End With
        Next i
        
        '--- end of file ------------------------------------------------------------------
        
        Echo "done reading " & loc(ff)
        Echo "file size is " & LOF(ff)
        Echo ""
        
        .loaded = True
    End With
    
    'close file
    Close #ff
    
    'success
    LoadFhxRig = True
    Exit Function
errorhandler:
    MsgBox "LoadFhxRig" & vbLf & err.description, vbCritical
End Function


'fill treeview
Public Sub FillTreeFhxRig(ByRef tree As MSComctlLib.TreeView)
    With fhxrig
        On Error GoTo errhandler
        If Not .loaded Then Exit Sub
        
        Dim n As MSComctlLib.node
        
        'file root
        Dim rootname As String
        rootname = "fhx_rig"
        Set n = tree.Nodes.Add(, tvwChild, rootname, GetFileName(.filename), "file")
        n.Expanded = True
        n.tag = 0
        
        'version
        tree.Nodes.Add rootname, tvwChild, "filever", "Version: " & .head.version, "prop"
        
        'number of nodes
        tree.Nodes.Add rootname, tvwChild, "nodes", "Nodes: " & .nodenum, "prop"
        
        'loop nodes
        Dim i As Long
        For i = 0 To .nodenum - 1
            With .node(i)
                
                Dim name As String
                name = "Node " & i
                
                If .parent = -1 Then
                    'root
                    Set n = tree.Nodes.Add(rootname, tvwChild, "node" & i, name, "lod")
                Else
                    'child
                    Set n = tree.Nodes.Add("node" & .parent, tvwChild, "node" & i, name, "lod")
                End If
                
                'add node
                n.Expanded = True
                n.tag = i
                
            End With
        Next i
        
    End With
    Exit Sub
errhandler:
    MsgBox "FillTreeFhxRig" & vbLf & err.description, vbCritical
End Sub


'draws rig
Public Sub DrawFhxRig()
Dim i As Long
    With fhxrig
        If Not .loaded Then Exit Sub
        
        glColor3f 1, 1, 0
        StartAAPoint 5
        StartAALine 1.3
        
        For i = 0 To .nodenum - 1
            
            'draw dot
            glBegin GL_POINTS
                'glVertex3fv .node(i).pos.x
                glVertex3fv .node(i).mat.m(12)
            glEnd
            
            'draw line
            If .node(i).parent > -1 Then
                glBegin GL_LINES
                    'glVertex3fv .node(i).pos.x
                    'glVertex3fv .node(.node(i).parent).pos.x
                    
                    glVertex3fv .node(i).mat.m(12)
                    glVertex3fv .node(.node(i).parent).mat.m(12)
                glEnd
            End If
            
        Next i
        
        EndAALine
        EndAAPoint
        
    End With
End Sub


Public Sub UnloadFhxRig()
    With fhxrig
        .loaded = False
        .filename = ""
        
        .nodenum = 0
        Erase .node()
    End With
End Sub

