Attribute VB_Name = "OBJ_Import"
Option Explicit

Private Type objface
    v1 As Integer
    v2 As Integer
    v3 As Integer
    
    t1 As Integer
    t2 As Integer
    t3 As Integer
    
    n1 As Integer
    n2 As Integer
    n3 As Integer
End Type

Private Type objmesh
    name As String
    facenum As Long
    face() As objface
End Type

Private Type objfile
    vertnum As Long
    texcnum As Long
    normnum As Long
    
    vert() As float3
    texc() As float2
    norm() As float3
    
    groupnum As Long
    group() As objmesh
    
    'internal
    filename As String
    loaded As Boolean
End Type

Public myobj As objfile


Private Sub ParseLine(ByVal ln As String)
    With myobj
        
        'removed extra spaces
        ln = Trim$(ln)
        ln = Replace(ln, "  ", " ")
        
        'exit if empty line
        If Len(ln) = 0 Then Exit Sub
        
        'split to tokens
        Dim str() As String
        str = Split(ln, " ")
        
        'process instructions
        Select Case str(0)
        Case "v"
            .vertnum = .vertnum + 1
            ReDim Preserve .vert(0 To .vertnum - 1)
            .vert(.vertnum - 1).x = val(str(1))
            .vert(.vertnum - 1).y = val(str(2))
            .vert(.vertnum - 1).z = val(str(3))
        Case "vt"
            .texcnum = .texcnum + 1
            ReDim Preserve .texc(0 To .texcnum - 1)
            .texc(.texcnum - 1).x = val(str(1))
            .texc(.texcnum - 1).y = val(str(2))
        Case "vn"
            .normnum = .normnum + 1
            ReDim Preserve .norm(0 To .normnum - 1)
            .norm(.normnum - 1).x = val(str(1))
            .norm(.normnum - 1).y = val(str(2))
            .norm(.normnum - 1).z = val(str(3))
        Case "g"
            .groupnum = .groupnum + 1
            ReDim Preserve .group(0 To .groupnum - 1)
            .group(.groupnum - 1).name = str(1)
        Case "f"
            
            'create group if none defined
            If .groupnum = 0 Then
                .groupnum = 1
                ReDim .group(0 To .groupnum - 1)
            End If
            
            'add face
            Dim tstr() As String
            With .group(.groupnum - 1)
                .facenum = .facenum + 1
                ReDim Preserve .face(0 To .facenum - 1)
                
                tstr = Split(str(1), "/")
                .face(.facenum - 1).v1 = val(tstr(0)) - 1
                .face(.facenum - 1).t1 = val(tstr(1)) - 1
                .face(.facenum - 1).n1 = val(tstr(2)) - 1
                
                tstr = Split(str(2), "/")
                .face(.facenum - 1).v2 = val(tstr(0)) - 1
                .face(.facenum - 1).t2 = val(tstr(1)) - 1
                .face(.facenum - 1).n2 = val(tstr(2)) - 1
                
                tstr = Split(str(3), "/")
                .face(.facenum - 1).v3 = val(tstr(0)) - 1
                .face(.facenum - 1).t3 = val(tstr(1)) - 1
                .face(.facenum - 1).n3 = val(tstr(2)) - 1
                
                'quads
                If UBound(str) > 3 Then
                    .facenum = .facenum + 1
                    ReDim Preserve .face(0 To .facenum - 1)
                    
                    tstr = Split(str(1), "/")
                    .face(.facenum - 1).v1 = val(tstr(0)) - 1
                    .face(.facenum - 1).t1 = val(tstr(1)) - 1
                    .face(.facenum - 1).n1 = val(tstr(2)) - 1
                    
                    tstr = Split(str(3), "/")
                    .face(.facenum - 1).v2 = val(tstr(0)) - 1
                    .face(.facenum - 1).t2 = val(tstr(1)) - 1
                    .face(.facenum - 1).n2 = val(tstr(2)) - 1
                    
                    tstr = Split(str(4), "/")
                    .face(.facenum - 1).v3 = val(tstr(0)) - 1
                    .face(.facenum - 1).t3 = val(tstr(1)) - 1
                    .face(.facenum - 1).n3 = val(tstr(2)) - 1
                End If
                
            End With
        End Select

    End With
End Sub


Public Function LoadOBJ(ByRef filename As String) As Boolean
Dim i As Long
Dim ln As String
Dim linenum As Long
    On Error GoTo errorhandler
    
    UnloadObj
    
    Dim ff As Integer
    ff = FreeFile
    Open filename For Input As #ff
    
    Echo ""
    
    With myobj
        .loaded = False
        .filename = filename
        
        Do Until EOF(ff)
            linenum = linenum + 1
            
            Line Input #ff, ln
            
            'split lines by unix linebreak
            Dim lnarr() As String
            lnarr = Split(ln, vbLf)
            
            For i = LBound(lnarr()) To UBound(lnarr())
                ParseLine lnarr(i)
            Next i
        Loop
        Close #ff
        
        'print stats
        Echo "vertnum: " & .vertnum
        Echo "texcnum: " & .texcnum
        Echo "normnum: " & .texcnum
        Echo "groupnum: " & .groupnum
        Echo ""
        
        'print groups
        For i = 0 To .groupnum - 1
            Echo "group " & i
            Echo " name: " & .group(i).name
            Echo " facenum: " & .group(i).facenum
            Echo ""
        Next i
        
        .loaded = True
    End With
    
    LoadOBJ = True
    Exit Function
errorhandler:
    MsgBox "LoadOBJ" & vbLf & err.Description & vbLf & "On line " & linenum, vbCritical
End Function


Public Sub DrawObj()
    On Error GoTo errorhandler
    
    With myobj
        If Not .loaded Then Exit Sub
        
        'draw faces
        If view_poly Then
            
            'draw solid
            If view_lighting Then
                glEnable GL_LIGHTING
            End If
            If view_edges Or view_verts Then
                glPolygonOffset 1, 1
                glEnable GL_POLYGON_OFFSET_FILL
            End If
            glColor3f 0.75, 0.75, 0.75
            DrawObjSimple
            If view_edges Or view_verts Then
                glDisable GL_POLYGON_OFFSET_FILL
            End If
            If view_lighting Then
                glDisable GL_LIGHTING
            End If
            
            'draw edges
            If view_edges And Not view_wire Then
                glColor4f 1, 1, 1, 0.1
                StartAALine 1.3
                glPolygonMode GL_FRONT_AND_BACK, GL_LINE
                DrawObjSimple
                glPolygonMode GL_FRONT_AND_BACK, GL_FILL
                EndAALine
            End If
            
        End If
        
        'draw vertices
        If view_verts Then
            StartAAPoint 4
            glColor3f 1, 1, 1
            glVertexPointer 3, GL_FLOAT, 0, .vert(0).x
            glEnableClientState GL_VERTEX_ARRAY
            
            glDrawArrays GL_POINTS, 0, .vertnum
            
            glDisableClientState GL_VERTEX_ARRAY
            EndAALine
        End If
        
        'mesh bounds
        'If view_bounds Then
        '    StartAALine 1.3
        '    glColor3f 1, 1, 0
        '    DrawBox .min, .max
        '    EndAALine
        'End If
        
    End With
        
    Exit Sub
errorhandler:
    MsgBox "DrawMyObj" & vbLf & err.Description, vbCritical
    myobj.loaded = False
End Sub


Public Sub DrawObjSimple()
Dim i As Long
Dim j As Long
Dim k As Long
    With myobj
        glBegin GL_TRIANGLES
        For i = 0 To .groupnum - 1
            With .group(i)
                For j = 0 To .facenum - 1
                    glNormal3fv myobj.norm(.face(j).n3).x
                    glTexCoord2fv myobj.texc(.face(j).t3).x
                    glVertex3fv myobj.vert(.face(j).v3).x
                    
                    glNormal3fv myobj.norm(.face(j).n2).x
                    glTexCoord2fv myobj.texc(.face(j).t2).x
                    glVertex3fv myobj.vert(.face(j).v2).x
                    
                    glNormal3fv myobj.norm(.face(j).n1).x
                    glTexCoord2fv myobj.texc(.face(j).t1).x
                    glVertex3fv myobj.vert(.face(j).v1).x
                Next j
            End With
        Next i
        glEnd
    End With
End Sub


Public Sub UnloadObj()
    With myobj
        .loaded = False
        .filename = ""
        
        .vertnum = 0
        .texcnum = 0
        .normnum = 0
        .groupnum = 0
        
        Erase .vert()
        Erase .texc()
        Erase .norm()
        Erase .group()
    End With
End Sub

'fills treeview hierarchy
Public Sub FillTreeObj(ByRef tree As MSComctlLib.TreeView)
    Dim i As Long
    Dim j As Long
    With myobj
        If Not .loaded Then Exit Sub
        
        Dim n As MSComctlLib.node
        
        'root
        Dim rootname As String
        rootname = "obj_root"
        Set n = tree.Nodes.Add(, tvwChild, rootname, GetFileName(.filename), "file")
        n.Expanded = True
        
        tree.Nodes.Add rootname, tvwChild, rootname & "|vertnum", "Vertices: " & .vertnum, "prop"
        tree.Nodes.Add rootname, tvwChild, rootname & "|texcnum", "Texcoords: " & .texcnum, "prop"
        tree.Nodes.Add rootname, tvwChild, rootname & "|normnum", "Normals: " & .normnum, "prop"
        
        For i = 0 To .groupnum - 1
            With .group(i)
                
                'add geom
                Dim groupname As String
                groupname = "group" & i
                Set n = tree.Nodes.Add(rootname, tvwChild, groupname, "Group " & i, "lod")
                n.Expanded = True
                
                'add geom properties
                tree.Nodes.Add groupname, tvwChild, groupname & "|name", "Name: " & .name, "prop"
                tree.Nodes.Add groupname, tvwChild, groupname & "|trinum", "Faces: " & .facenum, "trinum"
                
            End With
        Next i
        
    End With
End Sub


