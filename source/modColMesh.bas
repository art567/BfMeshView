Attribute VB_Name = "BF2_Collisionmesh"
Option Explicit


'geom face (8 bytes)
Private Type colface
    v1 As Integer
    v2 As Integer
    v3 As Integer
    m As Integer
End Type

'ystruct (16 bytes)
Private Type ystruct
    u1 As Single
    u2 As Integer
    u3 As Integer
    u4 As Long
    u5 As Long
End Type

'geom
Private Type bf2collod
    'collider type
    coltype As Long        '0=projectile, 1=vehicle, 2=soldier, 3=AI
    
    'face data
    facenum As Long
    face() As colface
    
    'vertex data
    vertnum As Long
    vert() As float3
    
    'unknown
    vertid() As Integer
    
    'vertex bounds
    min As float3
    max As float3
    
    'unknown
    u7 As Byte
    
    'tree bounds
    bmin As float3
    bmax As float3
    
    'unknown
    ynum As Long
    ydata() As ystruct
    
    'unknown
    znum As Long
    zdata() As Integer
    
    'unknown
    anum As Long
    adata() As Long
    
    '!!!internal!!!
    norm() As float3
    badtri As Long
End Type

''collisionmesh geom (temp!!)
'Private Type bf2colgeom
'    type As Long         '0=parachute, 1=staticmesh, 2=most bundledmesh, 3=vehicle / NO! its number of groups!!
'
'    qlodnum As Long
'    qlod() As bf2colgeomlod
'
'    'LODs
'    lodnum As Long
'    lod() As bf2colgeomlod
'
'    'wreck LODs
'    xlodnum As Long
'    xlod() As bf2colgeomlod
'End Type

'sub
Private Type bf2colsub
    lodnum As Long
    lod() As bf2collod
End Type

'geom
Private Type bf2colgeom
    subgnum As Long
    subg() As bf2colsub 'holds variations of this geom
End Type

'collisionmesh file
Private Type bf2col
    
    'header
    u1 As Long          '0  ?
    ver As Long         '8  file format version
    
    'geoms
    geomnum As Long     '1  number of geoms
    geom() As bf2colgeom
    
    ''''internal
    loadok As Boolean
    drawok As Boolean
    filename As String
End Type

Public cmesh As bf2col


'loads collisionmesh from file
Public Function LoadBF2Col(ByRef filename As String) As Boolean
    On Error GoTo errorhandler
    
    'open file
    Dim ff As Integer
    ff = FreeFile
    Open filename For Binary Access Read Lock Write As #ff
    
Dim i As Long
Dim j As Long
Dim c As Long
    
    With cmesh
        .loadok = False
        .drawok = True
        .filename = filename
        
        '--- header ---------------------------------------------------------
        
        'unknown (4 bytes)
        Get #ff, , .u1
        Echo "u1: " & .u1
        
        'version (4 bytes)
        Get #ff, , .ver
        Echo "version: " & .ver
        
        'version warning
        Select Case .ver
        Case 8:
        Case 9:
        Case 10:
        Case Else
            MsgBox "File type not tested, may crash!", vbExclamation
        End Select
        
        'geomnum (4 bytes)
        Get #ff, , .geomnum
        Echo "geomnum: " & .geomnum
        Echo ""
        
        'loop through geoms
        If .geomnum > 0 Then
            ReDim .geom(0 To .geomnum - 1)
            For i = 0 To .geomnum - 1
                Echo "geom " & c & " start at " & loc(ff)
                
                BF2ReadColGeom ff, .geom(i)
                
                Echo "geom " & c & " end at " & loc(ff)
                Echo ""
            Next i
        End If
        
        '--- end of file ------------------------------------------------------------------
        
        Echo "done reading " & loc(ff)
        Echo "file size is " & LOF(ff)
        Echo ""
        
        If loc(ff) <> LOF(ff) Then
            MsgBox "File not loaded properly!", vbExclamation
        End If
        
        'internal
        .loadok = True
        .drawok = True
    End With
    
    'close file
    Close #ff
    
    GenColNormals
    CheckBF2ColMesh
    
    'success
    LoadBF2Col = True
    Exit Function
errorhandler:
    MsgBox "LoadBF2Col" & vbLf & err.description, vbCritical
    Echo ">>> error at " & loc(ff)
    Echo ">>> file size " & LOF(ff)
    GenColNormals
    GenColorTable
End Function


'reads geom
Private Sub BF2ReadColGeom(ByRef ff As Integer, ByRef geom As bf2colgeom)
    Dim i As Long
    With geom
        
        Get #ff, , .subgnum
        Echo " subnum: " & .subgnum
        
        'read subs
        If .subgnum > 0 Then
            
            Echo ""
            ReDim .subg(0 To .subgnum - 1)
            For i = 0 To .subgnum - 1
                Echo " sub " & i & " start at " & loc(ff)
                
                BF2ReadColSub ff, .subg(i)
                
                Echo " sub " & i & " end at " & loc(ff)
                Echo ""
            Next i
            
        End If
        
    End With
End Sub


'reads sub
Private Sub BF2ReadColSub(ByRef ff As Integer, ByRef subg As bf2colsub)
    Dim i As Long
    With subg
        
        'lodnum (4 bytes)
        Get #ff, , .lodnum
        Echo " lodnum: " & .lodnum
        
        'read geoms
        If .lodnum > 0 Then
            Echo ""
            ReDim .lod(0 To .lodnum - 1)
            For i = 0 To .lodnum - 1
                Echo " lod " & i & " start at " & loc(ff)
                
                BF2ReadColLod ff, .lod(i)
                
                Echo " lod " & i & " end at " & loc(ff)
                Echo ""
            Next i
        End If
        
    End With
End Sub


'read collider geom block
Private Sub BF2ReadColLod(ByRef ff As Integer, ByRef lod As bf2collod)
    With lod
        
        'coltype (4 bytes)
        If cmesh.ver >= 9 Then
            Get #ff, , .coltype
            Echo "  id: " & .coltype
        End If
        
        '--- faces ---
        
        'facenum (4 bytes)
        Get #ff, , .facenum
        Echo "  facenum: " & .facenum
        
        'faces (8 bytes * facenum)
        If .facenum > 0 Then
            ReDim .face(0 To .facenum - 1)
            Get #ff, , .face()
        End If
        
        '--- vertices ---
        
        'vertnum (4 bytes)
        Get #ff, , .vertnum
        Echo "  vertnum: " & .vertnum
        
        'vertices (12 bytes * vertnum)
        If .vertnum > 0 Then
            ReDim .vert(0 To .vertnum - 1)
            Get #ff, , .vert()
        End If
        
        'vertid (2 bytes * vertnum)
        If .vertnum > 0 Then
            ReDim .vertid(0 To .vertnum - 1)
            Get #ff, , .vertid()
        End If
        
        '--- bounds ---
        
        'bounds (24 bytes)
        Get #ff, , .min
        Get #ff, , .max
        
        '--- misc ---
        
        'unknown (1 byte)
        Get #ff, , .u7 'always 49??
        Echo "  u7: " & .u7
        
        '--- misc ---
        
        'bounds (24 bytes)
        Get #ff, , .bmin
        Get #ff, , .bmax
        
        '--- y block ---
        
        'ynum (4 bytes)
        Get #ff, , .ynum
        Echo "  ynum: " & .ynum
        
        'ydata (16 bytes * ynum)
        If .ynum > 0 Then
            Echo "  ydata start at " & loc(ff)
            
            ReDim .ydata(0 To .ynum - 1)
            Get #ff, , .ydata()
            
            Echo "  ydata end at " & loc(ff)
        End If
        
        '--- z block ---
        
        'znum != facenum
        'could be index to triangle?
        
        'znum (4 bytes)
        Get #ff, , .znum
        Echo "  znum: " & .znum
        
        'zdata (2 bytes * znum)
        If .znum > 0 Then
            Echo "  zdata start at " & loc(ff)
            
            ReDim .zdata(0 To .znum - 1)
            Get #ff, , .zdata()
            
            Echo "  zdata start at " & loc(ff)
        End If
        
        '--- a block ---
        
        If cmesh.ver >= 10 Then
            'anum (4 bytes)
            Get #ff, , .anum
            Echo "  anum: " & .anum
            
            'adata (4 bytes * anum)
            If .anum > 0 Then
                Echo "  adata start at " & loc(ff)
                
                ReDim .adata(0 To .anum - 1)
                Get #ff, , .adata()
                
                Echo "  adata start at " & loc(ff)
            End If
        End If
        
    End With
End Sub


'counts number of degenerate triangles
Private Sub CheckBF2ColLod(ByRef lod As bf2collod)
Dim i As Long
Dim v1 As float3
Dim v2 As float3
Dim v3 As float3
Dim a1 As Single
Dim a2 As Single
Dim a3 As Single
Dim err As Boolean
    
    With lod
        .badtri = 0
        
        For i = 0 To .facenum - 1
            
            err = False
            
            v1 = .vert(.face(i).v1)
            v2 = .vert(.face(i).v2)
            v3 = .vert(.face(i).v3)
            
            a1 = AngleBetweenVectors(SubFloat3(v1, v2), SubFloat3(v1, v3))
            a2 = AngleBetweenVectors(SubFloat3(v2, v3), SubFloat3(v2, v1))
            a3 = AngleBetweenVectors(SubFloat3(v3, v1), SubFloat3(v3, v2))
            
            'DEGENERATEFACEANGLE
            Const badangle As Single = 0.1
            
            If a1 < badangle Then err = True
            If a2 < badangle Then err = True
            If a3 < badangle Then err = True
            
            If err Then
                .badtri = .badtri + 1
            End If
        Next i
        
    End With
    
End Sub


'checks for errors
Public Sub CheckBF2ColMesh()
Dim i As Long
Dim j As Long
Dim k As Long
    With cmesh
        If Not .loadok Then Exit Sub
        
        For i = 0 To .geomnum - 1
            With .geom(i)
                For j = 0 To .subgnum - 1
                    With .subg(j)
                        For k = 0 To .lodnum - 1
                            CheckBF2ColLod .lod(k)
                        Next k
                    End With
                Next j
            End With
        Next i
        
    End With
End Sub


'fill treeview
Public Sub FillTreeColMesh(ByRef tree As MSComctlLib.TreeView)
    On Error GoTo errhandler
    
    Dim n As MSComctlLib.node
    Dim tag As Long
    
    With cmesh
        'If Not .loadok Then Exit Sub
        If Not .drawok Then Exit Sub 'temp!!!
        
        'file root
        Dim rootname As String
        rootname = "bf2col_root"
        tag = MakeTag(0, 0, 0)
        Set n = tree.Nodes.Add(, , rootname, GetFileName(.filename), "file")
        n.tag = tag
        n.Expanded = True
        
        'version leaf
        Set n = tree.Nodes.Add(rootname, tvwChild, rootname & "|ver", "Version: " & .ver, "prop")
        n.tag = tag
        
        'loop geoms
        Dim i As Long
        For i = 0 To .geomnum - 1
            
            'add geom node
            Dim geomname As String
            geomname = "geom" & i
            tag = MakeTag(i, 0, 0)
            Set n = tree.Nodes.Add(rootname, tvwChild, geomname, "Geom " & i, "geom")
            n.tag = tag
            If .geom(i).subgnum > 0 Then n.Expanded = True
            
            'loop subs
            Dim j As Long
            For j = 0 To .geom(i).subgnum - 1
                
                'add sub node
                Dim subname As String
                subname = geomname & "|sub" & j
                tag = MakeTag(i, j, 0)
                Set n = tree.Nodes.Add(geomname, tvwChild, subname, "Sub " & j, "geom")
                n.tag = tag
                If .geom(i).subg(j).lodnum > 0 Then n.Expanded = True
                
                'loop subs
                Dim k As Long
                For k = 0 To .geom(i).subg(j).lodnum - 1
                    With .geom(i).subg(j).lod(k)
                        
                        Dim lodicon As String
                        If .badtri > 0 Then
                            lodicon = "badlod"
                        Else
                            lodicon = "lod"
                        End If
                        
                        'add lod node
                        Dim coltypestr As String
                        coltypestr = GetColTypeString(.coltype)
                        
                        Dim lodname As String
                        lodname = subname & "|lod" & k
                        tag = MakeTag(i, j, k)
                        Set n = tree.Nodes.Add(subname, tvwChild, lodname, "Col " & .coltype & " (" & coltypestr & ")", lodicon)
                        n.tag = tag
                        
                        'add facenum leaf
                        Set n = tree.Nodes.Add(lodname, tvwChild, lodname & "|trinum", "Faces: " & .facenum, "trinum")
                        n.tag = tag
                        
                        'add vertnum leaf
                        Set n = tree.Nodes.Add(lodname, tvwChild, lodname & "|vertnum", "Vertices: " & .vertnum, "prop")
                        n.tag = tag
                        
                     End With
                Next k
                
            Next j
            
        Next i
        
    End With
    Exit Sub
errhandler:
    MsgBox "FillTreeColMesh" & vbLf & err.description, vbCritical
End Sub


'returns string of collisionmesh LOD type
Private Function GetColTypeString(ByVal coltype As Long) As String
    Select Case coltype
    Case 0: GetColTypeString = "projectile"
    Case 1: GetColTypeString = "vehicle"
    Case 2: GetColTypeString = "soldier"
    Case 3: GetColTypeString = "AI"
    Case Else: GetColTypeString = "unknown"
    End Select
End Function


'draws collisionmesh
Public Sub DrawColMesh()
    If Not cmesh.drawok Then Exit Sub
    
    On Error GoTo errorhandler
    
    If selgeom < 0 Then Exit Sub
    If selgeom > cmesh.geomnum - 1 Then Exit Sub
    
    If selsub < 0 Then Exit Sub
    If selsub > cmesh.geom(selgeom).subgnum - 1 Then Exit Sub
    
    If sellod < 0 Then Exit Sub
    If sellod > cmesh.geom(selgeom).subg(selsub).lodnum - 1 Then Exit Sub
    
    DrawColLod cmesh.geom(selgeom).subg(selsub).lod(sellod)
    
    Exit Sub
errorhandler:
    MsgBox "DrawColMesh" & vbLf & err.description, vbCritical
    cmesh.drawok = False
End Sub


'draws collisionmesh geom
Private Sub DrawColLod(ByRef geom As bf2collod)
    On Error GoTo errorhandler
    With geom
        
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
            DrawColLodFaces geom
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
                DrawColLodFaces geom
                glPolygonMode GL_FRONT_AND_BACK, GL_FILL
                EndAALine
            End If
            
        End If
        
        'draw vertices
        If view_verts Then
            StartAAPoint 4
            glColor3f 1, 1, 1
            glVertexPointer 3, GL_FLOAT, 0, .vert(0).X
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
    MsgBox "DrawColMeshGeom" & vbLf & err.description, vbCritical
    cmesh.drawok = False
End Sub


'draws colored solid faces
Private Sub DrawColLodFaces(ByRef geom As bf2collod)
Dim i As Long
Dim v1 As Long
Dim v2 As Long
Dim v3 As Long
Dim m As Long
Dim cc As Long
    With geom
        glBegin GL_TRIANGLES
        For i = 0 To .facenum - 1
            
            v1 = .face(i).v3
            v2 = .face(i).v2
            v3 = .face(i).v1
            m = .face(i).m
            
            cc = Clamp(m, 0, maxcolors)
            glColor4fv colortable(cc).r
            
            glNormal3fv .norm(i).X
            
            glVertex3fv .vert(v1).X
            glVertex3fv .vert(v2).X
            glVertex3fv .vert(v3).X
        Next i
        glEnd
    End With
End Sub


'generates face normals
Private Sub GenColNormals()
    On Error GoTo errorhandler

    With cmesh
        Dim i As Long
        For i = 0 To .geomnum - 1
            With .geom(i)
                Dim j As Long
                For j = 0 To .subgnum - 1
                    With .subg(j)
                        Dim k As Long
                        For k = 0 To .lodnum - 1
                            GenColLodNormals .lod(k)
                        Next k
                    End With
                Next j
            End With
        Next i
    End With
    
    Exit Sub
errorhandler:
    'only raise error if file was loaded properly
    If cmesh.loadok Then
        MsgBox "GenColNormals" & vbLf & err.description, vbCritical
    End If
End Sub


'generates geom face normals
Private Sub GenColLodNormals(ByRef geom As bf2collod)
Dim i As Long
    With geom
        If .facenum = 0 Then Exit Sub
        ReDim .norm(0 To .facenum - 1)
        For i = 0 To .facenum - 1
            .norm(i) = GenNormal(.vert(.face(i).v1), .vert(.face(i).v2), .vert(.face(i).v3))
        Next i
    End With
End Sub


'clears collisionmesh
Public Sub UnloadBF2Col()
    With cmesh
        .loadok = False
        .drawok = False
    End With
End Sub


