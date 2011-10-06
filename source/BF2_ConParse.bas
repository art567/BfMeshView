Attribute VB_Name = "BF2_ConParse"
Option Explicit

Private Type bf2con_node
    parent As Long
    type As String
    name As String
    geometryPart As Long
    transform As matrix4
    
    'internal
    wtrans As matrix4
End Type

Private Type bf2con_file
    nodenum As Long
    node() As bf2con_node
    
    partnum As Long
    part() As Long
    
    'internal
    filename As String
    loaded As Boolean
End Type

Public bf2con As bf2con_file


'creates new node
Private Function bf2conCreateNode(ByRef t As String, ByRef name As String) As Long
    With bf2con
        
        'check if node already exists
        Dim i As Long
        For i = 0 To .nodenum - 1
            If .node(i).name = name Then
                bf2conCreateNode = i
                Exit Function
            End If
        Next i
        
        'add node
        .nodenum = .nodenum + 1
        ReDim Preserve .node(0 To .nodenum - 1)
        
        With .node(.nodenum - 1)
            .parent = -1
            .type = t
            .name = name
            .geometryPart = -1
            mat4identity .transform
        End With
        
        bf2conCreateNode = .nodenum - 1
    End With
End Function


'loads BF2 console script
Public Sub LoadCon(ByRef filename As String)
    On Error GoTo errhandler
    
    'check if file exists
    If Not FileExist(filename) Then
        Exit Sub
    End If
    
    'open file
    Dim ff As Integer
    ff = FreeFile
    Open filename For Input As #ff
    
    Dim linenum As Long
    Dim ln As String
    Dim str() As String
    Dim skip As Boolean
    
    Dim cnode As Long 'last created object
    Dim tnode As Long 'last created template (child of object)
    
    With bf2con
        .loaded = True
        .filename = filename
        
        Do Until EOF(ff)
            linenum = linenum + 1
            Line Input #ff, ln
            
            'remove whitespaces
            ln = Trim$(ln)
            
            'skip comments
            skip = False
            If Len(ln) = 0 Then skip = True
            If LCase(Left$(ln, 3)) = "rem" Then skip = True
            
            'process
            If Not skip Then
                str() = Split(ln, " ")
                Select Case str(0)
                    
                Case "ObjectTemplate.create"
                    cnode = bf2conCreateNode(str(1), str(2))
                    
                Case "ObjectTemplate.addTemplate"
                    tnode = bf2conCreateNode("<unknown>", str(1))
                    .node(tnode).parent = cnode
                    
                Case "ObjectTemplate.geometryPart"
                    .node(cnode).geometryPart = val(str(1))
                    
                Case "ObjectTemplate.setPosition"
                    str = Split(str(1), "/")
                    Dim pos As float3
                    pos.X = val(str(0))
                    pos.Y = val(str(1))
                    pos.z = val(str(2))
                    mat4setpos .node(tnode).transform, pos
                    
                Case "ObjectTemplate.setRotation"
                    str = Split(str(1), "/")
                    Dim rot As float3
                    rot.X = val(str(1))
                    rot.Y = val(str(0))
                    rot.z = val(str(2))
                    mat4setrotYXZ .node(tnode).transform, rot
                    
                End Select
            End If
            
        Loop
    End With
    
    'close file
    Close ff
    
    'post processing
    Dim i As Long
    With bf2con
     
        'build lookup table
        
        'count max number of geometryParts
        Dim max As Long
        For i = 0 To .nodenum - 1
            If .node(i).geometryPart > max Then max = .node(i).geometryPart
        Next i
        
        'allocate lookup table
        If max > 0 Then
            .partnum = max + 1
            ReDim .part(0 To .partnum)
            
            'reset table
            For i = 0 To .partnum - 1
                .part(i) = 0
            Next i
        End If
        
        'assign table
        For i = 0 To .nodenum - 1
            Dim gp As Long
            gp = .node(i).geometryPart
            If gp > -1 Then
               .part(gp) = i
            End If
        Next i
        
        'Echo "partnum: " & .partnum
        'For i = 0 To .partnum - 1
        '    Echo "part[" & i & "]: " & .part(i)
        'Next i
        
        'compute world space transformation matrices
        mat4identity .node(0).transform
        For i = 0 To .nodenum - 1
            If .node(i).parent > -1 Then
                .node(i).wtrans = mat4mult(.node(i).transform, .node(.node(i).parent).transform)
            Else
                .node(i).wtrans = .node(i).transform
            End If
        Next i
        
    End With
    
    'deform mesh
    BF2MeshDeform2
    
    'success
    On Error GoTo 0
    Exit Sub
    
    'error
errhandler:
    'MsgBox "LoadCon" & vbLf & err.description & vbLf & filename & " (" & linenum & ")", vbCritical
    Close ff
    On Error GoTo 0
End Sub


'unloads con
Public Sub UnloadCon()
    With bf2con
        .loaded = False
        .filename = ""
        .nodenum = 0
        .partnum = 0
        Erase .node()
        Erase .part()
    End With
End Sub


'draws node transform defined in CON
Public Sub DrawConNodes()
    With bf2con
        If Not .loaded Then Exit Sub
        
        If view_bonesys Then
            Const s = 0.01
            Dim min As float3
            Dim max As float3
            min.X = -s
            min.Y = -s
            min.z = -s
            max.X = s
            max.Y = s
            max.z = s
            
            glDisable GL_LIGHTING
            glDisable GL_TEXTURE_2D
            glDisable GL_DEPTH_TEST
            
            StartAALine 1.333
            Dim i As Long
            For i = 0 To .nodenum - 1
                With .node(i)
                    glPushMatrix
                        glMultMatrixf .wtrans.m(0)
                        
                        DrawPivot s * 2
                        
                        glColor3f 1, 1, 0
                        DrawBox min, max
                    glPopMatrix
                    
                    If .parent > -1 Then
                        glColor3f 0, 1, 1
                        glBegin GL_LINES
                            glVertex3fv .wtrans.m(12)
                            glVertex3fv bf2con.node(.parent).wtrans.m(12)
                        glEnd
                    End If
                End With
            Next i
            EndAALine
            
            glEnable GL_DEPTH_TEST
        End If
        
    End With
End Sub

'fill treeview
Public Sub FillTreeBF2Con(ByRef tree As MSComctlLib.TreeView)
    With bf2con
        On Error GoTo errhandler
        If Not .loaded Then Exit Sub
        
        Dim n As MSComctlLib.node
        
        'file root
        Dim rootname As String
        rootname = "bf2con"
        Set n = tree.Nodes.Add(, tvwChild, rootname, GetFileName(.filename), "file")
        n.Expanded = False
        'n.tag = 0
        
        'number of nodes
        tree.Nodes.Add rootname, tvwChild, "nodes", "Nodes: " & .nodenum, "prop"
        
        'loop nodes
        Dim i As Long
        For i = 0 To .nodenum - 1
            With .node(i)
                
                Dim name As String
                'name = "Node " & i & " (" & .name & ")" & .geometryPart
                name = .name & " (" & .geometryPart & ")"
                
                If .parent = -1 Then
                    'root
                    Set n = tree.Nodes.Add(rootname, tvwChild, "node" & i, name, "lod")
                Else
                    'child
                    Set n = tree.Nodes.Add("node" & .parent, tvwChild, "node" & i, name, "lod")
                End If
                
                'add node
                n.Expanded = True
                'n.tag = i
                
            End With
        Next i
        
    End With
    Exit Sub
errhandler:
    MsgBox "FillTreeBF2Con" & vbLf & err.description, vbCritical
End Sub

