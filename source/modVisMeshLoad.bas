Attribute VB_Name = "BF2_MeshLoad"
Option Explicit

Public vmesh As bf2mesh 'file structure


'loads mesh from file
Public Function LoadBF2Mesh(ByRef filename As String) As Boolean
    On Error GoTo errorhandler
    
    If Not FileExist(filename) Then
        MsgBox "File " & filename & " not found.", vbExclamation
        Exit Function
    End If
    
Dim i As Long
Dim j As Long
    
    'open file
    Dim ff As Integer
    ff = FreeFile
    Open filename For Binary Access Read Lock Write As #ff
    
    With vmesh
        
        'reset stuff
        .filename = filename
        .fileext = LCase(GetFileExt(filename))
        .isBundledMesh = (.fileext = "bundledmesh")
        .isSkinnedMesh = (.fileext = "skinnedmesh")
        '.isBFP4F = InStr(1, LCase(filename), "bfp4f")
        .drawok = True
        .loadok = False
        
        '--- header --------------------------------------------------------------------------------
        
        Echo "head start at " & loc(ff)
        Get #ff, , .head
        Echo " u1: " & .head.u1
        Echo " version: " & .head.version
        Echo " u3: " & .head.u3
        Echo " u4: " & .head.u4
        Echo " u5: " & .head.u5
        Echo "head end at " & loc(ff)
        Echo ""
        
        'unknown (1 byte)
        Get #ff, , .u1 'stupid little byte that misaligns the entire file!
        Echo "u1: " & .u1
        Echo ""
        'for BFP4F, the value is "1", so perhaps this is a version number as well
        If .u1 = 1 Then .isBFP4F = True
        
        '--- geom table ---------------------------------------------------------------------------
        
        Echo "geom table start at " & loc(ff)
        
        'geomnum (4 bytes)
        Get #ff, , .geomnum
        Echo " geomnum: " & .geomnum
        
        'geom table (4 bytes * groupnum)
        ReDim .geom(0 To .geomnum - 1)
        For i = 0 To .geomnum - 1
            
            'lodnum (4 bytes)
            Get #ff, , .geom(i).lodnum
            ReDim .geom(i).lod(0 To .geom(i).lodnum - 1)
            
            Echo "  lodnum: " & .geom(i).lodnum
        Next i
        
        Echo "geom table end at " & loc(ff)
        Echo ""
        
        '--- vertex attribute table -------------------------------------------------------------------------------
        
        Echo "attrib block at " & loc(ff)
        
        'vertattribnum (4 bytes)
        Get #ff, , .vertattribnum
        Echo " vertattribnum: " & .vertattribnum
        
        'vertex attributes
        ReDim .vertattrib(0 To .vertattribnum - 1)
        Get #ff, , .vertattrib()
        For i = 0 To .vertattribnum - 1
            Echo " attrib[" & i & "]: " & .vertattrib(i).flag & " " & _
                                          .vertattrib(i).offset & " " & _
                                          .vertattrib(i).vartype & " " & _
                                          .vertattrib(i).usage
        Next i
        
        Echo "attrib block end at " & loc(ff)
        Echo ""
        
        '--- vertices -----------------------------------------------------------------------------
        
        Echo "vertex block start at " & loc(ff)
        
        Get #ff, , .vertformat
        Get #ff, , .vertstride
        Get #ff, , .vertnum
        Echo " vertformat: " & .vertformat
        Echo " vertstride: " & .vertstride
        Echo " vertnum: " & .vertnum
        
        ReDim .vert(0 To ((.vertstride / .vertformat) * .vertnum) - 1)
        Get #ff, , .vert()
        
        Echo "vertex block end at " & loc(ff)
        Echo ""
        
        '--- indices ------------------------------------------------------------------------------
        
        Echo "index block start at " & loc(ff)
        
        Get #ff, , .indexnum
        Echo " indexnum: " & .indexnum
        ReDim .Index(0 To .indexnum - 1)
        Get #ff, , .Index()
        
        Echo "index block end at " & loc(ff)
        Echo ""
        
        '--- rigs -------------------------------------------------------------------------------
        
        'unknown (4 bytes)
        If Not .isSkinnedMesh Then
            Get #ff, , .u2 'always 8?
            Echo " u2: " & .u2
        End If
        
        'rigs/nodes
        Echo "nodes chunk start at " & loc(ff)
        For i = 0 To .geomnum - 1
            Echo " geom " & i & " start"
            For j = 0 To .geom(i).lodnum - 1
                Echo "  lod " & j & " start"
                ReadLodNodeTable ff, .geom(i).lod(j)
                Echo "  lod " & j & "  end"
            Next j
            Echo " geom " & i & "  end"
            Echo ""
        Next i
        Echo "nodes chunk end at " & loc(ff)
        Echo ""
        
        '--- triangles ------------------------------------------------------------------------------
        
        Echo "geom block start at " & loc(ff)
        
        For i = 0 To .geomnum - 1
            For j = 0 To .geom(i).lodnum - 1
                Echo " mesh " & j & " start at " & loc(ff)
                ReadGeomLod ff, .geom(i).lod(j)
                Echo " mesh " & j & " end at " & loc(ff)
            Next j
        Next i
        
        Echo "geom block end at " & loc(ff)
        Echo ""
        
        '--- end of file -------------------------------------------------------------------------
        
        Echo "done reading " & loc(ff)
        Echo "file size is " & LOF(ff)
        Echo ""
        
        .loadok = True
        .drawok = True
    End With
    
    'close file
    Close #ff
    
    'generate some useful stuff
    GenVertInfo
    
    'reset node transforms
    nodetransformnum = 40
    For i = 0 To 40 - 1
        mat4identity nodetransform(i)
    Next i
    
    'auto-load con
    If opt_loadcon Then
        With vmesh
            If .isBundledMesh Then
                
                'make con file title (e.g. "bedford.con")
                Dim conTitle As String
                conTitle = GetNameFromFileName(.filename) & ".con"
                
                'try parent folder
                Dim conName As String
                conName = GetFilePath(.filename) & "..\" & conTitle
                If FileExist(conName) Then
                    LoadCon conName
                Else
                    'try current folder
                    conName = GetFilePath(.filename) & "\" & conTitle
                    If FileExist(conName) Then
                        LoadCon conName
                    End If
                End If
                
            End If
        End With
    End If
    
    'detect UV channels
    With vmesh
        .uvnum = 0
        For i = 0 To .vertattribnum - 1
            If Not .vertattrib(i).flag = 255 Then
                If .vertattrib(i).vartype = 1 Then
                    .uvnum = .uvnum + 1
                End If
            End If
        Next i
    End With
    
    LoadBF2Mesh = True
    Exit Function
errorhandler:
    MsgBox "LoadBF2Mesh" & vbLf & err.description, vbCritical
    Echo ">>> error at " & loc(ff)
    Echo ">>> filesize " & LOF(ff)
    Close ff
End Function


'reads lod node table
Private Sub ReadLodNodeTable(ByRef ff As Integer, ByRef lod As bf2lod)
Dim i As Long
Dim j As Long
    With lod
        
        Echo ">>> " & loc(ff)
        
        'bounds (24 bytes)
        Get #ff, , .min
        Get #ff, , .max
        
        'unknown (12 bytes)
        If vmesh.head.version <= 6 Then 'version 4 and 6
            Get #ff, , .pivot
        End If
        
        If vmesh.isSkinnedMesh Then
            
            'rignum (4 bytes)
            Get #ff, , .rignum
            Echo "   rignum: " & .rignum
            
            'read rigs
            If .rignum > 0 Then
                ReDim .rig(0 To .rignum - 1)
                For i = 0 To .rignum - 1
                    With .rig(i)
                        Echo "   rig block " & i & " start at " & loc(ff)
                        
                        'bonenum (4 bytes)
                        Get #ff, , .bonenum
                        Echo "   bonenum: " & .bonenum
                        
                        'bones (68 bytes * bonenum)
                        If .bonenum > 0 Then
                            ReDim .bone(0 To .bonenum - 1)
                            For j = 0 To .bonenum - 1
                            
                                'bone id (4 bytes)
                                Get #ff, , .bone(j).id
                                
                                'bone transform (64 bytes)
                                Get #ff, , .bone(j).matrix
                                
                                Echo "    boneid[" & j & "]: " & .bone(j).id
                            Next j
                        End If
                        
                        Echo "   rig block " & i & " end at " & loc(ff)
                    End With
                    
                Next i
            End If
            
        Else
            
            'nodenum (4 bytes)
            Get #ff, , .nodenum
            Echo "   nodenum: " & .nodenum
            
            'node matrices (64 bytes * nodenum)
            If Not vmesh.isBundledMesh Then
                Echo "   node data"
                
                If .nodenum > 0 Then
                    ReDim .node(0 To .nodenum - 1)
                    For i = 0 To .nodenum - 1
                        Get #ff, , .node(i)
                    Next i
                End If
            End If
            
            'node matrices (BFP4F variant)
            If vmesh.isBundledMesh And vmesh.isBFP4F Then
                Echo "   node data"
                
                If .nodenum > 0 Then
                    ReDim .node(0 To .nodenum - 1)
                    For i = 0 To .nodenum - 1
                    
                        'matrix (64 bytes)
                        Get #ff, , .node(i)
                        
                        'name length (4 bytes)
                        Dim namelen As Long
                        Get #ff, , namelen
                        
                        'name (includes zero terminator)
                        Dim name() As Byte
                        ReDim name(0 To namelen - 1)
                        Get #ff, , name()
                        
                    Next i
                End If
            End If
            
        End If
        
    End With
End Sub


'reads lod material chunk
Private Sub ReadLodMat(ByRef ff As Integer, ByRef mat As bf2mat)
Dim i As Long
    With mat
        
        'alpha flag (4 bytes)
        If Not vmesh.isSkinnedMesh Then
            Get #ff, , .alphamode
            Echo "   alphamode: " & .alphamode
        End If
        
        'fx filename
        .fxfile = ReadString(ff)
        Echo "   fxfile: " & .fxfile
        
        'material name
        .technique = ReadString(ff)
        Echo "   matname: " & .technique
        
        'mapnum (4 bytes)
        Get #ff, , .mapnum
        Echo "   mapnum: " & .mapnum
        
        'maps (? bytes)
        If .mapnum > 0 Then
            ReDim .map(0 To .mapnum - 1)
            
            'mapnames
            For i = 0 To .mapnum - 1
                .map(i) = ReadString(ff)
                Echo "    " & .map(i)
            Next i
        End If
        
        'geometry info
        Get #ff, , .vstart
        Get #ff, , .istart
        Get #ff, , .inum
        Get #ff, , .vnum
        Echo "   vstart: " & .vstart
        Echo "   istart: " & .istart
        Echo "   inum: " & .inum
        Echo "   vnum: " & .vnum
        
        'unknown
        Get #ff, , .u4
        Get #ff, , .u5
        'note: filled garbage for BFP4F
        
        'bounds (24 bytes)
        If Not vmesh.isSkinnedMesh Then
            If vmesh.head.version = 11 Then
                Get #ff, , .mmin
                Get #ff, , .mmax
            End If
        End If
        
        '--- internal --------------------------------------
        If .mapnum > 0 Then
            ReDim .texmapid(0 To .mapnum - 1)
            ReDim .mapuvid(0 To .mapnum - 1)
        End If
        
    End With
End Sub


'reads geom lod chunk
Private Sub ReadGeomLod(ByRef ff As Integer, ByRef mesh As bf2lod)
Dim i As Long
    With mesh
        
        'internal: reset polycount
        .polycount = 0
        
        'matnum (4 bytes)
        Get #ff, , .matnum
        Echo "  matnum: " & .matnum
        
        'materials (? bytes)
        ReDim .mat(0 To .matnum - 1)
        For i = 0 To .matnum - 1
            Echo "  mat " & i & " start at " & loc(ff)
            ReadLodMat ff, .mat(i)
            Echo "  mat " & i & " end at " & loc(ff)
            
            'internal: increment polycount
            .polycount = .polycount + .mat(i).inum / 3
        Next i
        
    End With
End Sub


'reads string
Private Function ReadString(ByRef ff As Integer) As String
Dim num As Long
Dim chars() As Byte
    Get #ff, , num
    
    If num = 0 Then Exit Function
    
    ReDim chars(0 To num - 1)
    Get #ff, , chars()
    
    ReadString = SafeString(chars, num)
End Function


'clears file data
Public Sub UnloadBF2Mesh()
    With vmesh
        'we're lazy today, flip some bits and leave deallocation to garbage collector
        .loadok = False
        .drawok = False
        
        Erase .vertattrib()
        Erase .vert()
        Erase .Index()
        Erase .geom()
        
        'internal
        Erase .vertinfo()
        Erase .vertsel()
        Erase .vertflag()
        
        .hasSkinVerts = False
        Erase .skinvert()
        Erase .skinnorm()
    End With
End Sub


Public Function MakeKey(ByVal geo As Long, ByVal lod As Long, ByVal mat As Long, ByVal tex As Long) As String
    MakeKey = "@|" & geo & "|" & lod & "|" & mat & "|" & tex
End Function


'generates vertex info
Public Sub GenVertInfo()
    With vmesh
        
        'internal stuff
        .xstride = .vertstride / 4
        ReDim .vertinfo(0 To .vertnum - 1)
        ReDim .vertsel(0 To .vertnum - 1)
        ReDim .vertflag(0 To .vertnum - 1)
        
        'generate info
        Dim g As Long
        For g = 0 To .geomnum - 1
            With .geom(g)
                Dim L As Long
                For L = 0 To .lodnum - 1
                    With .lod(L)
                        Dim m As Long
                        For m = 0 To .matnum - 1
                            With .mat(m)
                                Dim v As Long
                                For v = .vstart To .vstart + .vnum - 1
                                    vmesh.vertinfo(v).geom = g
                                    vmesh.vertinfo(v).lod = L
                                    vmesh.vertinfo(v).mat = m
                                    vmesh.vertinfo(v).sel = 0
                                Next v
                            End With
                        Next m
                    End With
                Next L
            End With
        Next g
        
        'clear selection/flags
        Dim i As Long
        For i = 0 To .vertnum - 1
            .vertsel(i) = 0
            .vertflag(i) = 0
        Next i
        
    End With
End Sub


'fills treeview hierarchy
Public Sub FillTreeVisMesh(ByRef tree As MSComctlLib.TreeView)
Dim i As Long
Dim j As Long
Dim k As Long
Dim m As Long
Dim n As MSComctlLib.node
    
    With vmesh
        If Not .loadok Then Exit Sub
        
        'file root
        Dim rootname As String
        rootname = "bf2mesh_root"
        Set n = tree.Nodes.Add(, , rootname, GetFileName(.filename), "file")
        n.Expanded = True
        n.tag = 0
        
        'file version tag
        tree.Nodes.Add rootname, tvwChild, "filever", "Version: " & .head.version, "prop"
        
        'vertex stride
        tree.Nodes.Add rootname, tvwChild, "vertstride", "Vertex Stride: " & .vertstride & " bytes", "prop"
        
        'vertex number
        tree.Nodes.Add rootname, tvwChild, "vertnum", "Vertices: " & .vertnum, "prop"
        
        'vertex attributes
        Set n = tree.Nodes.Add(rootname, tvwChild, "vertattribnum", "Vertex Attributes: " & .vertattribnum - 1, "prop")
        For i = 0 To .vertattribnum - 1
            If Not .vertattrib(i).flag = 255 Then
                
                Dim typename As String
                Select Case .vertattrib(i).vartype
                Case 0: typename = "FLOAT1"
                Case 1: typename = "FLOAT2"
                Case 2: typename = "FLOAT3"
                Case 3: typename = "FLOAT4"
                Case 4: typename = "D3DCOLOR"
                Case 5: typename = "UBYTE4"
                Case 6: typename = "SHORT2"
                Case 7: typename = "SHORT4"
                Case Else: typename = "unknown type"
                End Select
                
                Dim attribname As String
                Select Case .vertattrib(i).usage
                Case 0: attribname = "position"
                Case 1: attribname = "blend weight"
                Case 2: attribname = "blend indices"
                Case 3: attribname = "normal"
                Case 5: attribname = "uv1"
                Case 6: attribname = "tangent"
                Case 261: attribname = "uv2"
                Case 517: attribname = "uv3"
                Case 773: attribname = "uv4"
                Case 1029: attribname = "uv5"
                Case Else: attribname = "unknown"
                End Select
                
                tree.Nodes.Add n, tvwChild, "vattrib" & i, attribname & " (" & typename & ")", "prop"
            End If
        Next i
        
        Dim nodename As String
        
        'loop geoms
        For i = 0 To .geomnum - 1
            With .geom(i)
                nodename = MakeKey(i, -1, -1, -1)
                
                'add geom node
                Dim geo As MSComctlLib.node
                Set geo = tree.Nodes.Add(rootname, tvwChild, nodename, "Geom " & i, "geom")
                geo.tag = -1
                If vmesh.geomnum < 4 Then geo.Expanded = True
                
                'loop lods
                For j = 0 To .lodnum - 1
                    With .lod(j)
                        nodename = MakeKey(i, j, -1, -1)
                        
                        'add lod node
                        Dim lod As MSComctlLib.node
                        Set lod = tree.Nodes.Add(geo, tvwChild, nodename, "Lod " & j, "lod")
                        lod.Expanded = True
                        lod.tag = -1
                        
                        Set n = tree.Nodes.Add(lod, tvwChild, nodename & "|trinum", "Polygons: " & .polycount, "trinum")
                        n.tag = -1
                        
                        If vmesh.isSkinnedMesh Then
                            Set n = tree.Nodes.Add(lod, tvwChild, nodename & "|rignum", "Rigs: " & .rignum, "prop")
                            n.tag = -1
                        Else
                            Set n = tree.Nodes.Add(lod, tvwChild, nodename & "|nodenum", "Nodes: " & .nodenum, "prop")
                            n.tag = -1
                        End If
                        
                        'materials
                        For k = 0 To .matnum - 1
                            With .mat(k)
                                nodename = MakeKey(i, j, k, -1)
                                
                                'add material node
                                Dim mat As MSComctlLib.node
                                Set mat = tree.Nodes.Add(lod, tvwChild, nodename, "Material " & k, "mat")
                                mat.tag = -1
                                
                                Dim alphastr As String
                                If .alphamode = 0 Then alphastr = "None"
                                If .alphamode = 1 Then alphastr = "Alpha Blend"
                                If .alphamode = 2 Then alphastr = "Alpha Test"
                                
                                Set n = tree.Nodes.Add(mat, tvwChild, nodename & "|alpha", "Transparency: " & alphastr, "prop")
                                n.tag = -1
                                
                                Set n = tree.Nodes.Add(mat, tvwChild, nodename & "|shader", "Shader: " & .fxfile, "shader")
                                n.tag = -1
                                
                                Set n = tree.Nodes.Add(mat, tvwChild, nodename & "|technique", "Technique: " & .technique, "shader")
                                n.tag = -1
                                
                                Set n = tree.Nodes.Add(mat, tvwChild, nodename & "|trinum", "Polygons: " & (.inum / 3), "trinum")
                                n.tag = -1
                                
                                'textures
                                For m = 0 To .mapnum - 1
                                    nodename = MakeKey(i, j, k, m)
                                    
                                    Dim iconstr As String
                                    If .texmapid(m) = 0 Then
                                        iconstr = "texmissing"
                                    Else
                                        iconstr = "tex"
                                    End If
                                    
                                    'add texture node
                                    Set n = tree.Nodes.Add(mat, tvwChild, nodename, .map(m), iconstr)
                                    n.tag = -1
                                    
                                Next m
                            End With
                        Next k
                    End With
                Next j
            End With
        Next i
        
    End With
End Sub
