Attribute VB_Name = "a_DDS"
Option Explicit


Private Type ddscolorkey '8 bytes
    dw1 As Long
    dw2 As Long
End Type

Private Type ddspixelformat '32 bytes
    dwSize As Long
    dwFlags As Long
    dwFourCC As String * 4
    u1 As Long
    u2 As Long
    u3 As Long
    u4 As Long
    u5 As Long
End Type

Private Type ddscaps '16 bytes
    dwCaps As Long
    dwCaps2 As Long
    dwCaps3 As Long
    dwCaps4 As Long
End Type

Private Type ddsheader '124 bytes
    dwSize As Long
    dwFlags As Long
    dwWidth As Long
    dwHeight As Long
    dwLinearSize As Long
    dwBackBufferCount As Long
    dwMipMapCount As Long
    dwAlphaBitDepth As Long
    dwReserved As Long
    lpSurface As Long
    ddckCKDestOverlay As ddscolorkey
    ddckCKDestBlt As ddscolorkey
    ddckCKSrcOverlay As ddscolorkey
    ddckCKSrcBlt As ddscolorkey
    ddpfPixelFormat As ddspixelformat
    ddscaps As ddscaps
    dwTextureStage As Long
End Type


'loads texture from DDS file
Public Function LoadDDS(ByVal filename As String) As GLuint
    On Error GoTo errorhandler
    
    'bugfix?
    glActiveTexture GL_TEXTURE0
    glClientActiveTexture GL_TEXTURE0
    
    Echo "DDS: " & Chr(34) & filename & Chr(34)
    
    ''If GLEW_EXT_texture_compression_s3tc Then
    'If glewIsSupported("GL_EXT_texture_compression_s3tc") Then
    '    Echo "extension is supported"
    '    MsgBox "extension is supported"
    'Else
    '    MsgBox "S3TC texture compression (DXT5) extension not supported.", vbExclamation
    '    Exit Function
    'End If
    
    'check if file exists
    If Not FileExist(filename) Then
        MsgBox "File " & Chr(34) & filename & Chr(34) & " not found.", vbExclamation
        Exit Function
    End If
    
    'open file
    Dim ff As Integer
    ff = FreeFile
    Open filename For Binary Access Read Lock Write As #ff
    
    'read fourcc
    Dim fourcc As String * 4
    Get #ff, , fourcc
    'todo: check fourcc
    
    'read surface descriptor
    Dim ddsd As ddsheader
    Get #ff, , ddsd
    
    '...
    Dim fcode As String
    fcode = Left(ddsd.ddpfPixelFormat.dwFourCC, 4)
    If Left(fcode, 3) <> "DXT" Then fcode = ""
    'MsgBox "[" & SafeStr(fcode) & "]"
    Echo ">>> [" & fcode & "]"
    
    'determine DSS type
    Dim comp As Boolean
    Dim format As GLint
    Dim frags As Long
    Dim blocksize As Long
    Dim mipmapfactor As Long
    Select Case fcode
    Case "" 'uncompressed
        format = GL_BGRA
        frags = 4
        'frags = ddsd.ddpfPixelFormat.u1 / 8
        blocksize = 0
        mipmapfactor = 1
        comp = False
    Case "DXT1" 'DXT1 compression ratio is 8:1
        format = GL_COMPRESSED_RGBA_S3TC_DXT1_EXT
        frags = 4
        blocksize = 8
        mipmapfactor = 2
        comp = True
    Case "DXT3" 'DXT3 compression ratio is 4:1
        format = GL_COMPRESSED_RGBA_S3TC_DXT3_EXT
        frags = 4
        blocksize = 16
        mipmapfactor = 4
        comp = True
    Case "DXT5" 'DXT5 compression ratio is 4:1
        format = GL_COMPRESSED_RGBA_S3TC_DXT5_EXT
        frags = 4
        blocksize = 16
        mipmapfactor = 4
        comp = True
    Case Else
        Echo "DSS format not supported."
        Close ff
        Exit Function
    End Select
    
    Echo "format: " & ddsd.ddpfPixelFormat.dwFourCC
    
    'copy info
    Dim mapwidth As Long
    Dim mapheight As Long
    Dim mipmapnum As Long
    mapwidth = ddsd.dwWidth
    mapheight = ddsd.dwHeight
    If ddsd.dwMipMapCount = 0 Then
        mipmapnum = 1
    Else
        mipmapnum = ddsd.dwMipMapCount
    End If
    Echo "width: " & mapwidth
    Echo "height: " & mapheight
    Echo "mipmaps: " & mipmapnum
    
    'determine data size
    Dim datasize As Long
    If mipmapnum > 1 Then
        datasize = (mapwidth * mapheight * frags) * mipmapfactor
    Else
        datasize = (mapwidth * mapheight * frags)
    End If
    
    'read data
    Dim data() As Byte
    ReDim data(0 To datasize - 1)
    Get #ff, , data()
    
    'close file
    Close #ff
    
    'create texture
    Dim texname As GLuint
    glGenTextures 1, texname
    glBindTexture GL_TEXTURE_2D, texname
    
    'wrapping
    glTexParameteri GL_TEXTURE_2D, GL_TEXTURE_WRAP_S, GL_REPEAT
    glTexParameteri GL_TEXTURE_2D, GL_TEXTURE_WRAP_T, GL_REPEAT
    
    'filtering
    If mipmapnum > 1 Then
        glTexParameteri GL_TEXTURE_2D, GL_TEXTURE_MIN_FILTER, GL_LINEAR_MIPMAP_LINEAR
    Else
        glTexParameteri GL_TEXTURE_2D, GL_TEXTURE_MIN_FILTER, GL_LINEAR
    End If
    glTexParameteri GL_TEXTURE_2D, GL_TEXTURE_MAG_FILTER, GL_LINEAR
    
    'set max mipmap level, not all DDS texture may have all mipmaps
    If mipmapnum > 1 Then
        glTexParameteri GL_TEXTURE_2D, GL_TEXTURE_MAX_LEVEL, mipmapnum - 1
    End If
    
    'anisotropic filtering
    glTexParameterf GL_TEXTURE_2D, GL_TEXTURE_MAX_ANISOTROPY_EXT, maxaniso
    
    If comp Then
        
        'upload mipmaps
        Dim offset As Long
        Dim w As Long
        Dim h As Long
        Dim s As Long
        Dim i As Long
        w = mapwidth
        h = mapheight
        
        Dim dw As Long
        Dim dh As Long
        
        offset = 0
        For i = 0 To mipmapnum - 1
            
            dw = w / 4
            dh = h / 4
            If w < 4 Then dw = 1
            If h < 4 Then dh = 1
            s = dw * dh * blocksize
            
            'Echo ">>> mipmap " & i & ": " & w & "x" & h
            
            glCompressedTexImage2D GL_TEXTURE_2D, i, format, h, w, 0, s, ByVal VarPtr(data(offset))
            
            'check for errors
            Dim r As Long
            r = glGetError()
            If r <> GL_NO_ERROR Then
                Echo ">>> error on mipmap " & i & ": " & r
            End If
            
            offset = offset + s
            w = w / 2
            h = h / 2
            If w < 1 Then w = 1
            If h < 1 Then h = 1
        Next i
    Else
        glTexParameteri GL_TEXTURE_2D, GL_GENERATE_MIPMAP, GL_TRUE
        glTexImage2D GL_TEXTURE_2D, 0, frags, mapwidth, mapheight, 0, format, GL_UNSIGNED_BYTE, ByVal VarPtr(data(0))
        'gluBuild2DMipmaps GL_TEXTURE_2D, intformat, header.width, header.height, format, GL_UNSIGNED_BYTE, ByVal VarPtr(data(0))
    End If
    
    'clean up
    Erase data()
    
    'success
    Echo ""
    LoadDDS = texname
    Exit Function
errorhandler:
    MsgBox "LoadDDS" & vbLf & err.description, vbCritical
End Function
