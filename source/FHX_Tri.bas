Attribute VB_Name = "FHX_Tri"
Option Explicit

'collider geom face
Private Type fhxtriface
    v1 As Integer
    v2 As Integer
    v3 As Integer
End Type

'collider file
Private Type fhxtri_type
    'file header (16 bytes)
    head As fileheader
    
    'surface table
    surfacenum As Long
    surface() As String
    
    'geometry info
    vertnum As Long
    facenum As Long
    
    'vertex bounds
    min As float3
    max As float3
    
    'geometry data
    vert() As float3
    face() As fhxtriface
    facenorm() As float3
    faceid() As Byte
    
    'internal
    filename As String
    loadok As Boolean
    drawok As Boolean
End Type

Public fhxtri As fhxtri_type


'reads string from file
Private Function FhxTriReadString(ByRef ff As Integer)
Dim num As Byte
Dim chars() As Byte
    Get #ff, , num
    
    If num = 0 Then Exit Function
    
    ReDim chars(0 To num - 1)
    Get #ff, , chars()
    
    FhxTriReadString = SafeString(chars, num)
End Function


'reads FHX collider file
Public Function LoadFhxTri(ByVal filename As String)
    On Error GoTo errhandler
    
Dim i As Long
Dim j As Long
    
    'open file
    Dim ff As Integer
    ff = FreeFile
    Open filename For Binary Access Read Lock Write As #ff
    
    With fhxtri
        .loadok = False
        .filename = filename
        
        'read header
        Get #ff, , .head
        Echo "file format version: " & .head.version
        Echo "size: " & .head.size & "/" & LOF(ff)
        Echo "offset: " & .head.offset
        Echo ""
                
        'read surface table size
        Get #ff, , .surfacenum
        Echo " surfacenum: " & .surfacenum
        
        'read surface table
        ReDim .surface(0 To .surfacenum - 1)
        For j = 0 To .surfacenum - 1
            .surface(j) = FhxTriReadString(ff)
            Echo " surface[" & j & "]: " & .surface(j)
        Next j
        
        Get #ff, , .vertnum
        Get #ff, , .facenum
        Echo " vertnum: " & .vertnum
        Echo " facenum: " & .facenum
        
        'read bounds
        Echo " bounds at " & loc(ff)
        Get #ff, , .min
        Get #ff, , .max
        
        'read vertices
        If .vertnum > 0 Then
            Echo " vertices at " & loc(ff)
            ReDim .vert(0 To .vertnum - 1)
            Get #ff, , .vert()
        End If
        
        'read faces
        If .facenum > 0 Then
            Echo " faces at " & loc(ff)
            ReDim .face(0 To .facenum - 1)
            ReDim .facenorm(0 To .facenum - 1)
            ReDim .faceid(0 To .facenum - 1)
            
            For j = 0 To .facenum - 1
                Get #ff, , .face(j) 'todo: try removing loop (not sure byte alignment is a problem)
            Next j
            Get #ff, , .facenorm()
            Get #ff, , .faceid()
        End If
        
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
    LoadFhxTri = True
    Exit Function
    
    'error handler
errhandler:
    MsgBox "fhxLoadCol" & vbLf & err.Description, vbCritical
End Function


'draws geometry
Public Sub DrawFhxTri()
    If Not fhxtri.loadok Then Exit Sub
    If Not fhxtri.drawok Then Exit Sub
    On Error GoTo errorhandler
    
    Dim i As Long
    Dim j As Long
    
    Dim v As float3
    Dim n As float3
    Dim t1 As float3
    Dim t2 As float3
    
    With fhxtri
        
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
            DrawFhxTriFaces
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
                DrawFhxTriFaces
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
        
        'draw bounds
        If view_bounds Then
            StartAALine 1.3
            glColor3f 1, 1, 0
            DrawBox .min, .max
            EndAALine
        End If
        
    End With
    
    Exit Sub
errorhandler:
    MsgBox "fhxDrawCol()" & err.Description, vbCritical
    fhxtri.drawok = False
End Sub


'draws trimesh faces
Private Sub DrawFhxTriFaces()
Dim i As Long
Dim c As Long
    With fhxtri
        glBegin GL_TRIANGLES
        For i = 0 To .facenum - 1
            
            c = Clamp(.faceid(i), 0, maxcolors)
            glColor4fv colortable(c).r
            
            glNormal3fv .facenorm(i).x
            glVertex3fv .vert(.face(i).v1).x
            glVertex3fv .vert(.face(i).v2).x
            glVertex3fv .vert(.face(i).v3).x
        Next i
        glEnd
    End With
End Sub


'unloads collider data
Public Sub UnloadFhxTri()
    With fhxtri
        'clear internal
        .loadok = False
        .drawok = False
        .filename = ""
        
        'clear data
        .vertnum = 0
        .facenum = 0
        .surfacenum = 0
        Erase .surface()
        'erase .surfacename()
        Erase .vert()
        Erase .face()
        Erase .facenorm()
        Erase .faceid()
    End With
End Sub


'fills treeview hierarchy
Public Sub FillTreeFhxTri(ByRef tree As MSComctlLib.TreeView)
    Dim i As Long
    Dim j As Long
    With fhxtri
        If Not .loadok Then Exit Sub
        
        Dim n As MSComctlLib.node
        
        'file root
        Dim rootname As String
        rootname = "fhx_tri"
        Set n = tree.Nodes.Add(, , rootname, GetFileName(.filename), "file")
        n.Expanded = True
        n.tag = 0
        
        'version
        tree.Nodes.Add rootname, tvwChild, rootname & "|ver", "Version: " & .head.version, "prop"
        
        'properties
        tree.Nodes.Add rootname, tvwChild, rootname & "|surfnum", "Surfaces: " & .surfacenum, "prop"
        tree.Nodes.Add rootname, tvwChild, rootname & "|trinum", "Faces: " & .facenum, "trinum"
        tree.Nodes.Add rootname, tvwChild, rootname & "|vertnum", "Vertices: " & .vertnum, "prop"
            
    End With
End Sub

