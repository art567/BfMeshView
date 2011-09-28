Attribute VB_Name = "BF1942_Skin"
Option Explicit

Private Type skn_vert '31 bytes
    pos As float3
    
End Type

Private Type skn_file
    version As Long
    
    'geometry
    vertnum As Long
    
    'bones
    bonenum As Long
    bonename() As String * 8
    
    'internal
    filename As String
    loaded As Boolean
End Type

Public bfskin As skn_file


'loads rig from file
Public Function LoadSkin(ByVal filename As String) As Boolean
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
    LoadSkin = True
    Exit Function
errorhandler:
    MsgBox "LoadSkin" & vbLf & err.description, vbCritical
End Function


'fill treeview
Public Sub FillTreeSkin(ByRef tree As MSComctlLib.TreeView)
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
    MsgBox "FillTreeSkin" & vbLf & err.description, vbCritical
End Sub


'draws BF1942 skin
Public Sub DrawSkin()
Dim i As Long
    With fhxrig
        If Not .loaded Then Exit Sub
        
        glColor3f 1, 1, 0
        StartAAPoint 5
        StartAALine 1.3
        
        '
        
        EndAALine
        EndAAPoint
        
    End With
End Sub


'unloads BF1942 skin
Public Sub UnloadSkin()
    With fhxrig
        .loaded = False
        .filename = ""
        
        
    End With
End Sub

