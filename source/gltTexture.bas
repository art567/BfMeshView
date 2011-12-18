Attribute VB_Name = "a_TGA"
Option Explicit

Public Type tga_header
    offset As Byte
    colortype As Byte
    imagetype As Byte
    colormapstart As Integer
    colormaplength As Integer
    colormapbits As Byte
    xstart As Integer
    ystart As Integer
    width As Integer
    height As Integer
    bits As Byte
    flip As Byte
End Type


'loads texture from TGA file
Public Function LoadTGA(ByRef filename As String) As GLuint
Dim header As tga_header
Dim size As Long
Dim frags As Long
Dim data() As Byte
Dim handle As GLuint
Dim format As GLenum
Dim intformat As GLint
Dim errstr As String
    On Error GoTo errorhandler
    
    'bugfix?
    glActiveTexture GL_TEXTURE0
    glClientActiveTexture GL_TEXTURE0
    
    'check if file exists
    If Not FileExist(filename) Then
        MsgBox "File " & Chr(34) & filename & Chr(34) & " not found.", vbExclamation
        Exit Function
    End If
    
    'open file
    Dim ff As Integer
    ff = FreeFile
    Open filename For Binary As ff
    
    'get header data
    Get #ff, , header
    If Not IsPowerOfTwo(header.width) Then
        errstr = "Image width dimension not supported."
        GoTo errorhandler
    End If
    If Not IsPowerOfTwo(header.height) Then
        errstr = "Image height dimension not supported."
        GoTo errorhandler
    End If
    
    'create texture
    Select Case header.imagetype
    Case 1 'paletted
        
        'todo: add palette support
        'todo: detect grayscale palette
        
        If header.colormaplength = 0 Or header.colormaplength > 256 Then
            errstr = "No image palette data, or number of palette entries not supported."
            GoTo errorhandler
        End If
        
        'compute size
        size = CLng(header.width) * CLng(header.height)
        frags = 1
        
        'get pixel data
        ReDim data(0 To (size * frags) - 1)
        Get #ff, 1 + 18 + header.offset + (CLng(header.colormaplength) * (CLng(header.colormapbits) / 8)), data()
        
        'texture format
        intformat = GL_LUMINANCE
        format = GL_LUMINANCE
        
    Case 2 'rgb
        
        'compute size
        size = CLng(header.width) * CLng(header.height)
        frags = CLng(header.bits) / 8
        
        'get pixel data
        ReDim data(0 To (size * frags) - 1)
        Get #ff, 1 + 18 + header.offset, data()
        
        'texture format
        Select Case frags
        Case 3
            intformat = 3 'GL_RGB
            format = GL_BGR
        Case 4
            intformat = 4 'GL_RGBA
            format = GL_BGRA
        End Select
        
    Case 3 'grayscale
        
        'compute size
        size = CLng(header.width) * CLng(header.height)
        frags = 1
        
        'get pixel data
        ReDim data(0 To (size * frags) - 1)
        Get #ff, 1 + 18 + header.offset, data()
        
        'texture format
        intformat = GL_LUMINANCE
        format = GL_LUMINANCE
        
    Case Else
        errstr = "No image data, or image data format not supported."
        GoTo errorhandler
    End Select
    
    'close file
    Close ff
    
    'flip texture
    Dim y As Long
    Dim w As Long
    Dim h As Long
    Dim rowsize As Long
    Dim srcoff As Long
    Dim dstoff As Long
    w = header.width
    h = header.height
    rowsize = w * frags
    ReDim fdata(0 To (size * frags) - 1) As Byte
    For y = 0 To h - 1
        srcoff = y * rowsize
        dstoff = ((h - 1) - y) * rowsize
        CopyMem VarPtr(fdata(dstoff)), VarPtr(data(srcoff)), rowsize
    Next y
    
    'create texture
    glGenTextures 1, handle
    glBindTexture GL_TEXTURE_2D, handle
    glTexParameteri GL_TEXTURE_2D, GL_TEXTURE_WRAP_S, GL_REPEAT
    glTexParameteri GL_TEXTURE_2D, GL_TEXTURE_WRAP_T, GL_REPEAT
    glTexParameteri GL_TEXTURE_2D, GL_TEXTURE_MIN_FILTER, GL_LINEAR_MIPMAP_LINEAR
    glTexParameteri GL_TEXTURE_2D, GL_TEXTURE_MAG_FILTER, GL_LINEAR
    glTexParameterf GL_TEXTURE_2D, GL_TEXTURE_MAX_ANISOTROPY_EXT, maxaniso
    'glTexParameteri GL_TEXTURE_2D, GL_GENERATE_MIPMAP, GL_TRUE
    gluBuild2DMipmaps GL_TEXTURE_2D, intformat, header.width, header.height, format, GL_UNSIGNED_BYTE, ByVal VarPtr(fdata(0))
    
    'clean up
    Erase fdata()
    Erase data()
    
    'return handle
    LoadTGA = handle
    
    On Error GoTo 0
    Exit Function
errorhandler:
    If Len(errstr) > 0 Then
        MsgBox errstr, vbCritical, "LoadTGA"
    Else
        MsgBox err.description, vbCritical, "LoadTGA"
    End If
    On Error GoTo 0
End Function


'write TGA
Public Function WriteTGA(ByVal filename As String, ByVal w As Long, ByVal h As Long, ByVal bits As Long, ByRef data() As Byte)
    On Error GoTo errorhandler

    'open file
    Dim ff As Integer
    ff = FreeFile
    Open filename For Binary As #ff
    
    'create header
    Dim head As tga_header
    With head
        .offset = 0
        If bits = 8 Then
            'indexed grayscale
            .colortype = 1
            .imagetype = 1
            .colormapstart = 0
            .colormaplength = 256
            .colormapbits = 24
        Else
            'rgb/rgba
            .colortype = 0
            .imagetype = 2
        End If
        .xstart = 0
        .ystart = 0
        .width = w
        .height = h
        .bits = bits
        .flip = 0
    End With
    
    'write header
    With head
        Put #ff, , .offset
        Put #ff, , .colortype
        Put #ff, , .imagetype
        Put #ff, , .colormapstart
        Put #ff, , .colormaplength
        Put #ff, , .colormapbits
        Put #ff, , .xstart
        Put #ff, , .ystart
        Put #ff, , .width
        Put #ff, , .height
        Put #ff, , .bits
        Put #ff, , .flip
    End With
    
    'write palette
    If bits = 8 Then
        Dim i As Long
        Dim c As Byte
        For i = 0 To 255
            c = i
            Put #ff, , c
            Put #ff, , c
            Put #ff, , c
        Next i
    End If
    
    'write pixels
    Put #ff, , data()
    
    'close file
    Close #ff
    
    WriteTGA = True
    Exit Function
errorhandler:
    MsgBox "WriteTGA" & vbLf & err.description, vbCritical
    Close #ff
End Function


'writes 32-bit TGA
Public Function WriteTGA32(ByVal filename As String, ByVal w As Long, ByVal h As Long, ByRef data() As bgra)
    On Error GoTo errorhandler
    
    'open file
    Dim ff As Integer
    ff = FreeFile
    Open filename For Binary As #ff
    
    'create header
    Dim head As tga_header
    With head
        .offset = 0
        .colortype = 0
        .imagetype = 2
        .xstart = 0
        .ystart = 0
        .width = w
        .height = h
        .bits = 32
        .flip = 0
    End With
    
    'write header
    With head
        Put #ff, , .offset
        Put #ff, , .colortype
        Put #ff, , .imagetype
        Put #ff, , .colormapstart
        Put #ff, , .colormaplength
        Put #ff, , .colormapbits
        Put #ff, , .xstart
        Put #ff, , .ystart
        Put #ff, , .width
        Put #ff, , .height
        Put #ff, , .bits
        Put #ff, , .flip
    End With
    
    'write pixels
    Put #ff, , data()
    
    'close file
    Close #ff
    
    WriteTGA32 = True
    Exit Function
errorhandler:
    MsgBox "WriteTGA32" & vbLf & err.description, vbCritical
    Close #ff
End Function

