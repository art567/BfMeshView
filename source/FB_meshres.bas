Attribute VB_Name = "FB_meshres"
Option Explicit

'collider geom face
Private Type fbmeshface
    v1 As Integer
    v2 As Integer
    v3 As Integer
End Type

'collider file
Private Type fbmesh_type
    'file header (16 bytes)
    head As fileheader
    
    'geometry info
    vertnum As Long
    facenum As Long
    
    'vertex bounds
    min As float3
    max As float3
    
    'geometry data
    vert() As Single
    face() As fbmeshface
    
    'internal
    loadok As Boolean
    drawok As Boolean
    vertstride As Long
End Type

Public fbmesh As fbmesh_type


'reads FrostBite mesh file
Public Function LoadFbMesh(ByVal filename As String)
    On Error GoTo errhandler
    
Dim i As Long
Dim j As Long
    
    'open file
    Dim ff As Integer
    ff = FreeFile
    Open filename For Binary Access Read Lock Write As #ff
    
    With fbmesh
        
       ' '4 bytes
       ' Get #ff, , .u1 '0
       '
       ' '2 bytes
       ' Get #ff, , .u2 '0
       '
       ' '1 byte
       ' Get #ff, , .u3 '2, number of shaders?
       '
       ' For i = 0 To .u3 - 1
       '
        '    '.shadera = ReadFBmeshstring(ff)
        '
        '    '4 bytes
        '    Get #ff, , u4 '82
        '
        '    '4 bytes
        '    Get #ff, , u5 '188
        '
        '    '4 bytes
        '    Get #ff, , u6 '0
        '
        '    '4 bytes
        '    Get #ff, , u7 '0
        '
        '    '4 bytes
        '    Get #ff, , u8 '816
        '
        '    '2 bytes
        '    Get #ff, , u9 '0
        'Next i
        
        
        'read vertices
        Seek #ff, 147 + 1
        ReDim .vert(0 To (28 * 16) - 1)
        Get #ff, , .vert()
        
        'read faces
        Seek #ff, 1939 + 1
        ReDim .face(0 To 28 - 1)
        Get #ff, , .face()
        
        .vertstride = 0 '64 / 4
        
        'done
        Echo "done reading " & loc(ff)
        Echo "file size is " & LOF(ff)
        Echo ""
        
        .loadok = True
        .drawok = True
    End With
    
    'close file
    Close ff
    
    'success
    LoadFbMesh = True
    Exit Function
    
    'error handler
errhandler:
    MsgBox "fhxLoadCol" & vbLf & err.Description, vbCritical
End Function


'reads string from file
Private Function FbMeshReadString(ByRef ff As Integer)
Dim num As Byte
Dim chars() As Byte
    Get #ff, , num
    
    If num = 0 Then Exit Function
    
    ReDim chars(0 To num - 1)
    Get #ff, , chars()
    
    FbMeshReadString = SafeString(chars, num)
End Function


'draws geometry
Public Sub DrawFbMesh()
    If Not fbmesh.loadok Then Exit Sub
    If Not fbmesh.drawok Then Exit Sub
    On Error GoTo errorhandler
    
    Dim i As Long
    Dim j As Long
    
    Dim v As float3
    Dim n As float3
    Dim t1 As float3
    Dim t2 As float3
    
    With fbmesh
        
        'draw faces
        If view_poly Then
            
            'draw solid
            If view_lighting Then
                glEnable GL_LIGHTING
            End If
            'If view_edges Then
                glPolygonOffset 1, 1
                glEnable GL_POLYGON_OFFSET_FILL
            'End If
            glColor3f 0.75, 0.75, 0.75
            DrawFbMeshFaces
            'If view_edges Then
                glDisable GL_POLYGON_OFFSET_FILL
            'End If
            If view_lighting Then
                glDisable GL_LIGHTING
            End If
            
            'draw edges
            If view_edges And Not view_wire Then
                glColor4f 1, 1, 1, 0.1
                StartAALine 1.3
                glPolygonMode GL_FRONT_AND_BACK, GL_LINE
                DrawFbMeshFaces
                glPolygonMode GL_FRONT_AND_BACK, GL_FILL
                EndAALine
            End If
            
        End If
        
        'draw vertices
        If view_verts Then
            StartAAPoint 4
            glColor3f 1, 1, 1
            glVertexPointer 3, GL_FLOAT, .vertstride, .vert(0)
            glEnableClientState GL_VERTEX_ARRAY
            glDrawArrays GL_POINTS, 0, .vertnum
            glDisableClientState GL_VERTEX_ARRAY
            EndAALine
        End If
        
        'draw bounds
        'If view_bounds Then
        '    StartAALine 1.3
        '    glColor3f 1, 1, 0
        '    DrawBox .min, .max
        '    EndAALine
        'End If
        
    End With
    
    Exit Sub
errorhandler:
    MsgBox "fhxDrawCol()" & err.Description, vbCritical
    fhxtri.drawok = False
End Sub


'draws trimesh faces
Private Sub DrawFbMeshFaces()
Dim i As Long
    With fbmesh
        
        glBegin GL_TRIANGLES
        For i = 0 To .facenum - 1
            
            glVertex3fv .vert(.face(i).v1 * .vertstride)
            glVertex3fv .vert(.face(i).v2 * .vertstride)
            glVertex3fv .vert(.face(i).v3 * .vertstride)
            
        Next i
        glEnd
    End With
End Sub


'unloads collider data
Public Sub UnloadFbMesh()
    With fhxtri
        'clear internal
        .loadok = False
        .drawok = False
        
        
    End With
End Sub


'fills treeview hierarchy
Public Sub FillTreeFbMesh(ByRef tree As MSComctlLib.TreeView)
Dim i As Long
Dim j As Long
Dim n As MSComctlLib.node
Dim rootname As String
Dim geomname As String
    
    With fbmesh
        If Not .loadok Then Exit Sub
         
        ''root
        'rootname = "Collider"
        'Set n = tree.Nodes.Add(, , rootname, "Collider", "geom")
        'n.Expanded = True
        '
        ''version
        'tree.Nodes.Add rootname, tvwChild, rootname & "|ver", "Version: " & .head.version, "prop"
        '
        ''properties
        'tree.Nodes.Add rootname, tvwChild, rootname & "|surfnum", "Surfaces: " & .surfacenum, "prop"
        'tree.Nodes.Add rootname, tvwChild, rootname & "|trinum", "Faces: " & .facenum, "trinum"
        'tree.Nodes.Add rootname, tvwChild, rootname & "|vertnum", "Vertices: " & .vertnum, "prop"
            
    End With
End Sub


