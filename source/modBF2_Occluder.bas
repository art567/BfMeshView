Attribute VB_Name = "BF2_Occluder"
Option Explicit

Private Type occ_plane
    v1 As Long
    v2 As Long
    v3 As Long
    v4 As Long
End Type

Private Type occ_group
    planenum As Long
    plane() As occ_plane
    vertnum As Long
    vert() As float3
End Type

Private Type occ_file
    groupnum As Long
    group() As occ_group
    
    'internal
    filename As String
    loadok As Boolean
End Type

Private bf2occ As occ_file


'loads occluder from file
Public Function LoadOccluder(ByRef filename As String) As Boolean
    'On Error GoTo errhandler
    
    Dim ln As String
    Dim str() As String
    
    Dim planeindex As Long
    Dim vertindex As Long
    
    With bf2occ
        .loadok = False
        .filename = filename
    
        'open file
        Dim ff As Integer
        ff = FreeFile
        Open filename For Input As #ff
        
        Do Until EOF(ff)
            Line Input #ff, ln
            
            If Len(ln) > 0 Then
                If ln = "GROUP" Then
                    .groupnum = .groupnum + 1
                    ReDim Preserve .group(0 To .groupnum - 1)
                Else
                    If .groupnum > 0 Then
                        With .group(.groupnum - 1)
                            If .planenum = 0 Then
                                .planenum = val(ln)
                                ReDim .plane(0 To .planenum - 1)
                                planeindex = 0
                            Else
                                If planeindex < .planenum Then
                                    str() = Split(ln, " ")
                                    .plane(planeindex).v1 = val(str(0))
                                    .plane(planeindex).v2 = val(str(1))
                                    .plane(planeindex).v3 = val(str(2))
                                    .plane(planeindex).v4 = val(str(3))
                                    planeindex = planeindex + 1
                                Else
                                    If .vertnum = 0 Then
                                        .vertnum = val(ln)
                                        ReDim .vert(0 To .vertnum - 1)
                                        vertindex = 0
                                    Else
                                        str() = Split(ln, " ")
                                        .vert(vertindex).x = val(str(0))
                                        .vert(vertindex).y = val(str(1))
                                        .vert(vertindex).z = val(str(2))
                                        vertindex = vertindex + 1
                                    End If
                                End If
                            End If
                        End With
                    End If
                End If
                
            End If
            
        Loop
        Close ff
        
        .loadok = True
    End With
    
    LoadOccluder = True
    Exit Function
errhandler:
    MsgBox "LoadOccluder" & vbLf & err.Description, vbCritical
End Function


'draws occluder
Public Sub DrawOccluder()
    With bf2occ
        If Not .loadok Then Exit Sub
        
        Dim i As Long
        Dim j As Long
        
        glColor3f 0.9, 0.6, 0.1
        glDisable GL_LIGHTING
        glDisable GL_TEXTURE_2D
        
        glPolygonOffset 1, 1
        glEnable GL_POLYGON_OFFSET_FILL
        DrawPoly
        glDisable GL_POLYGON_OFFSET_FILL
        
        StartAALine 2
        glColor3f 1, 1, 1
        glPolygonMode GL_FRONT_AND_BACK, GL_LINE
        DrawPoly
        glPolygonMode GL_FRONT_AND_BACK, GL_FILL
        EndAALine
        
    End With
End Sub


'draws occluder polygons
Private Sub DrawPoly()
    With bf2occ
        Dim i As Long
        Dim j As Long
        For i = 0 To .groupnum - 1
            With .group(i)
                For j = 0 To .planenum - 1
                    glBegin GL_QUADS
                        glVertex3fv .vert(.plane(j).v1).x
                        glVertex3fv .vert(.plane(j).v2).x
                        glVertex3fv .vert(.plane(j).v3).x
                        glVertex3fv .vert(.plane(j).v4).x
                    glEnd
                Next j
            End With
        Next i
    End With
End Sub


'unloads occluder data
Public Sub UnloadOccluder()
    With bf2occ
        .loadok = False
        .filename = ""
        
        .groupnum = 0
        Erase .group()
    End With
End Sub


'fill treeview
Public Sub FillTreeOcc(ByRef tree As MSComctlLib.TreeView)
    On Error GoTo errhandler
    
    Dim n As MSComctlLib.node
    Dim tag As Long
    
    With bf2occ
        If Not .loadok Then Exit Sub
        
        'add root node
        Dim rootname As String
        rootname = "bf2_occ"
        tag = MakeTag(0, 0, 0)
        Set n = tree.Nodes.Add(, tvwChild, rootname, GetFileName(.filename), "file")
        n.tag = tag
        n.Expanded = True
        
        'version leaf
        'Set n = tree.Nodes.Add(rootname, tvwChild, rootname & "|ver", "Version: " & .ver, "prop")
        'n.tag = tag
        
        'loop geoms
        Dim i As Long
        For i = 0 To .groupnum - 1
            
            'add group node
            Dim groupname As String
            groupname = "geom" & i
            tag = MakeTag(i, 0, 0)
            Set n = tree.Nodes.Add(rootname, tvwChild, groupname, "Group " & i + 1, "geom")
            n.tag = tag
            n.Expanded = True
            
            With .group(i)
            
                'add planenum leaf
                Set n = tree.Nodes.Add(groupname, tvwChild, groupname & "|planenum", "Planes: " & .planenum, "trinum")
                n.tag = tag
                
                'add vertnum leaf
                Set n = tree.Nodes.Add(groupname, tvwChild, groupname & "|vertnum", "Vertices: " & .vertnum, "prop")
                n.tag = tag
            
            End With
            
        Next i
        
    End With
    Exit Sub
errhandler:
    MsgBox "FillTreeOcc" & vbLf & err.Description, vbCritical
End Sub

