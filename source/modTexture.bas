Attribute VB_Name = "a_Texture"
Option Explicit

Private Type byte4
    r As Byte
    g As Byte
    b As Byte
    a As Byte
End Type


'texture array
Private Type texmap_type
    filename As String
    origrelfilename As String
    filesize As Long
    tex As GLuint
End Type
Public texmapnum As Long
Public texmap() As texmap_type

Public dummytex As GLuint


'loads texture from file
Public Function LoadTexture(ByVal filename As String, ByVal origrelfilename As String) As Long
    
    'you wouldn't believe how badly Windows handles slashes
    filename = Replace(filename, "/", "\")
    
    'check to see if already loaded
    Dim i As Long
    For i = 1 To texmapnum
        If texmap(i).filename = filename Then
            LoadTexture = i
            Exit Function
        End If
    Next i
    
    SetStatus "info", "Loading " & filename
    
    'add new texture
    texmapnum = texmapnum + 1
    ReDim Preserve texmap(1 To texmapnum)
    
    'load texture file
    Dim ext As String
    ext = LCase(GetFileExt(filename))
    Select Case ext
    Case "tga": texmap(texmapnum).tex = LoadTGA(filename)
    Case "dds": texmap(texmapnum).tex = LoadDDS(filename)
    End Select
    texmap(texmapnum).filename = filename
    texmap(texmapnum).origrelfilename = origrelfilename
    texmap(texmapnum).filesize = GetFileSize(filename)
    
    'success
    LoadTexture = texmapnum
End Function


'returns filename string of texture to be loaded, or nothing if not found
Private Function BF2GetTexFileName(ByRef fname As String) As String
    
    'first try suffixed
    'Dim sfname As String
    'sfname = Replace(fname, ".dds", suffix(suffix_sel) & ".dds")
    'If FileExist(sfname) Then
    '    BF2GetTexFileName = sfname
    '    Exit Function
    'End If
   
    'try non-suffixed
    If FileExist(fname) Then
        BF2GetTexFileName = fname
        Exit Function
    End If
    
    'we ain't found shit
    BF2GetTexFileName = ""
End Function


'loads all mesh textures
Public Sub LoadMeshTextures()
    On Error GoTo errorhandler
    
Dim i As Long
Dim j As Long
Dim k As Long
Dim m As Long
Dim p As Long

Dim fname As String
Dim fpath As String
Dim mapfile As String
Dim filename As String
    
    'unload existing textures
    UnloadMeshTextures
    
    'BF2 mesh
    With vmesh
        If .loadok Then
            
            Dim meshfilepath As String
            meshfilepath = GetFilePath(vmesh.filename)
            
            For i = 0 To .geomnum - 1
                For j = 0 To .geom(i).lodnum - 1
                    For k = 0 To .geom(i).lod(j).matnum - 1
                        
                        'try to load all maps
                        For m = 0 To .geom(i).lod(j).mat(k).mapnum - 1
                            
                            'clear texmapid
                            .geom(i).lod(j).mat(k).texmapid(m) = 0
                            
                            'get map filename
                            mapfile = .geom(i).lod(j).mat(k).map(m)
                            
                            If Len(mapfile) > 0 Then
                                
                                'If suffix_sel > 0 Then
                                '    mapfile = replace(mapfile,
                                'end If
                                
                                'reset path
                                filename = ""
                                
                                If opt_uselocaltexpath Then
                                    
                                  '  'try mesh path first
                                  '  fname = meshfilepath & GetFileName(mapfile)
                                  '  If BF2TexFileExist(fname) Then filename = fname
                                    
                                  '  'try texture path
                                  '  If Len(filename) = 0 Then
                                  '      fname = meshfilepath & "..\Textures\" & GetFileName(mapfile)
                                  '      If BF2TexFileExist(fname) Then filename = fname
                                  '  End If
                                    
                                    'try mesh path first
                                    filename = BF2GetTexFileName(meshfilepath & GetFileName(mapfile))
                                    
                                    'try texture path
                                    If Len(filename) = 0 Then
                                        filename = BF2GetTexFileName(meshfilepath & "..\Textures\" & GetFileName(mapfile))
                                    End If
                                    
                                End If
                                
                                'try all of the texture folders
                                If Len(filename) = 0 Then
                                    fname = GetFileName(mapfile)
                                    For p = 1 To texpathnum
                                        If texpath(p).use Then
                                            
                                            'texpath+path+filename
                                            filename = BF2GetTexFileName(texpath(p).path & "\" & mapfile)
                                            If Len(filename) > 0 Then Exit For
                                            
                                            'texpath+filename
                                            If Len(filename) = 0 Then
                                                filename = BF2GetTexFileName(texpath(p).path & "\" & fname)
                                                If Len(filename) > 0 Then Exit For
                                            End If
                                            
                                        End If
                                    Next p
                                End If
                                
                                'load texture
                                If Len(filename) > 0 Then
                                    'Echo "trying to load " & filename
                                    
                                    'If j = 0 Then
                                    '    MsgBox .geom(i).lod(j).mat(k).texmapid(m)
                                    'End If
                                    
                                    .geom(i).lod(j).mat(k).texmapid(m) = LoadTexture(filename, mapfile)
                                    
                                    BuildShader .geom(i).lod(j).mat(k), vmesh.filename
                                    
                                Else
                                    Echo "Texture " & Chr(34) & mapfile & Chr(34) & " not loaded!" & vbCrLf
                                End If
                                
                            End If
                            
                        Next m
                    
                    Next k
                Next j
            Next i
            
        End If
    End With
    
    'bf1942 standard mesh
    With stdmesh
        If .loadok And stdshader.loaded Then
            
            For i = 1 To stdshader.subshader_num
                With stdshader.subshader(i)
                    For p = 1 To texpathnum
                        If texpath(p).use Then
                            
                            fname = texpath(p).path & "\" & .texture & ".dds"
                            If FileExist(fname) Then
                                .texmapid = LoadTexture(fname, "")
                            End If
                            
                            fname = texpath(p).path & "\" & .texture & ".tga"
                            If FileExist(fname) Then
                                .texmapid = LoadTexture(fname, "")
                            End If
                            
                        End If
                    Next p
                End With
            Next i
            
        End If
    End With
    
    
    'bf1942 tree mesh
    With treemesh
        If .loadok Then
            
            For i = 0 To .meshnum - 1
                With .mesh(i)
                    For j = 0 To .matnum - 1
                        With .mat(j)
                            For p = 1 To texpathnum
                                If texpath(p).use Then
                                    
                                    fname = texpath(p).path & "\" & .texname & ".dds"
                                    If FileExist(fname) Then
                                        .texmapid = LoadTexture(fname, "")
                                    End If
                                    
                                End If
                            Next p
                        End With
                    Next j
                End With
            Next i
            
        End If
    End With
    
    
    Exit Sub
errorhandler:
    MsgBox "LoadMeshTextures" & vbLf & err.description, vbCritical
End Sub


'unloads all mesh textures
Public Sub UnloadMeshTextures()
Dim i As Long
    For i = 1 To texmapnum
        texmap(i).filename = ""
        If texmap(i).tex <> 0 Then
            glDeleteTextures 1, texmap(i).tex
            texmap(i).tex = 0
        End If
    Next i
    Erase texmap()
    texmapnum = 0
End Sub


Public Sub BindTexture(ByVal id As Long)
    If id < 1 Or id > texmapnum Then
        UnbindTexture
        Exit Sub
    End If
    glBindTexture GL_TEXTURE_2D, texmap(id).tex
    glEnable GL_TEXTURE_2D
    glDisable GL_TEXTURE_CUBE_MAP
End Sub
Public Sub UnbindTexture()
    glBindTexture GL_TEXTURE_2D, 0
    glDisable GL_TEXTURE_2D
    glDisable GL_TEXTURE_CUBE_MAP
End Sub


'...
Public Function GetTextureMemory() As String
Dim total As Long
Dim i As Long
    For i = 1 To texmapnum
        total = total + texmap(i).filesize
    Next i
    GetTextureMemory = FormatFileSize(total)
End Function


'returns texture filename on disk
Public Function BF2GetTextureFilename(ByVal geo As Long, ByVal lod As Long, ByVal mat As Long, ByVal tex As Long) As String
    On Error GoTo errhandler
    
    If Not vmesh.loadok Then Exit Function
    If geo < 0 Then Exit Function
    If lod < 0 Then Exit Function
    If mat < 0 Then Exit Function
    If tex < 0 Then Exit Function
    
    If geo > vmesh.geomnum - 1 Then Exit Function
    If lod > vmesh.geom(geo).lodnum - 1 Then Exit Function
    If mat > vmesh.geom(geo).lod(lod).matnum - 1 Then Exit Function
    If tex > vmesh.geom(geo).lod(lod).mat(mat).mapnum - 1 Then Exit Function
    
    Dim texmapid As Long
    texmapid = vmesh.geom(geo).lod(lod).mat(mat).texmapid(tex)
    If texmapid > 0 Then
        BF2GetTextureFilename = texmap(texmapid).filename
    End If
    
    Exit Function
errhandler:
    MsgBox "BF2GetTextureFilename" & vbLf & err.description, vbCritical
End Function


Public Function GenTexture(ByVal r As Byte, ByVal g As Byte, ByVal b As Byte, ByVal a As Byte) As GLuint
    Dim tex As GLuint
    Dim pix(3) As Byte
    pix(0) = r
    pix(1) = g
    pix(2) = b
    pix(3) = a
    glGenTextures 1, tex
    glBindTexture GL_TEXTURE_2D, tex
    glTexParameteri GL_TEXTURE_2D, GL_TEXTURE_WRAP_S, GL_REPEAT
    glTexParameteri GL_TEXTURE_2D, GL_TEXTURE_WRAP_T, GL_REPEAT
    glTexParameteri GL_TEXTURE_2D, GL_TEXTURE_MIN_FILTER, GL_LINEAR
    glTexParameteri GL_TEXTURE_2D, GL_TEXTURE_MAG_FILTER, GL_LINEAR
    glTexImage2D GL_TEXTURE_2D, 0, GL_RGBA, 1, 1, 0, GL_RGBA, GL_UNSIGNED_BYTE, ByVal VarPtr(pix(0))
    glBindTexture GL_TEXTURE_2D, 0
    GenTexture = tex
End Function

'byte4 constructor
Public Function byte4(ByVal r As Byte, ByVal g As Byte, ByVal b As Byte, ByVal a As Byte) As byte4
    byte4.r = r
    byte4.g = g
    byte4.b = b
    byte4.a = a
End Function

Private Function IsEven(ByVal v As Long) As Boolean
    If v = 0 Then Exit Function
    IsEven = (val(v) Mod 2) = 0
End Function

'gen dummy
Public Sub GenDummyTex()
    Dim w As Long
    Dim h As Long
    w = 512
    h = 512
    Dim size As Long
    size = w * h
    Dim data() As byte4
    ReDim data(0 To size - 1)
    Dim red As byte4
    Dim yellow As byte4
    red = byte4(255, 0, 0, 127)
    yellow = byte4(255, 255, 0, 63)
    Dim x As Long
    Dim y As Long
    For y = 0 To h - 1
        For x = 0 To w - 1
            Dim i As Long
            i = x + w * y
            If IsEven(CLng(x \ 8) + CLng(y \ 8)) Then
                data(i) = red
            Else
                data(i) = yellow
            End If
        Next x
    Next y
    
    'create texture
    Dim handle As GLuint
    glGenTextures 1, handle
    glBindTexture GL_TEXTURE_2D, handle
    glTexParameteri GL_TEXTURE_2D, GL_TEXTURE_WRAP_S, GL_REPEAT
    glTexParameteri GL_TEXTURE_2D, GL_TEXTURE_WRAP_T, GL_REPEAT
    glTexParameteri GL_TEXTURE_2D, GL_TEXTURE_MIN_FILTER, GL_LINEAR_MIPMAP_LINEAR
    glTexParameteri GL_TEXTURE_2D, GL_TEXTURE_MAG_FILTER, GL_LINEAR
    glTexParameterf GL_TEXTURE_2D, GL_TEXTURE_MAX_ANISOTROPY_EXT, maxaniso
    glTexParameteri GL_TEXTURE_2D, GL_GENERATE_MIPMAP, GL_TRUE
    gluBuild2DMipmaps GL_TEXTURE_2D, GL_RGBA, w, h, GL_RGBA, GL_UNSIGNED_BYTE, ByVal VarPtr(data(0).r)
    'glTexImage2D GL_TEXTURE_2D, 0, GL_RGB, w, h, 0, GL_RGB, GL_UNSIGNED_BYTE, VarPtr(data(0))
    glBindTexture GL_TEXTURE_2D, 0
    
    Erase data()
    
    dummytex = handle
End Sub

