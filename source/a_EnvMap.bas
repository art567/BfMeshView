Attribute VB_Name = "a_EnvMap"
Option Explicit

Public Const envmapChan As Long = 8
Public envmapTex As GLuint


'loads envmap
Public Sub LoadEnvMap()
    If envmapTex <> 0 Then Exit Sub
    envmapTex = LoadTGAEnvMap(App.path & "\shaders\envmap.tga")
End Sub


'reloads envmap
Public Sub ReloadEnvMap()
    glDeleteTextures 1, envmapTex
    envmapTex = 0
    LoadEnvMap
End Sub


'loads envmap from TGA file
Public Function LoadTGAEnvMap(ByRef filename As String) As GLuint
    On Error GoTo errhandler
    
    'bugfix?
    glActiveTexture GL_TEXTURE0
    glClientActiveTexture GL_TEXTURE0
    
    'check if file exists
    If Not FileExist(filename) Then
    '    MsgBox "File " & Chr(34) & filename & Chr(34) & " not found.", vbExclamation
        Exit Function
    End If
    
    'open file
    Dim ff As Integer
    ff = FreeFile
    Open filename For Binary As ff
    
    'get header data
    Dim header As tga_header
    Get #ff, , header
    If Not IsPowerOfTwo(header.width) Then
        MsgBox "Image width dimension not supported.", vbExclamation
        Exit Function
    End If
    If Not IsPowerOfTwo(header.height / 6) Then
        MsgBox "Image height dimension not supported.", vbExclamation
        Exit Function
    End If
    
    'create texture
    If Not header.imagetype = 2 Then
        MsgBox "EnvMap format not supported.", vbExclamation
    End If
        
    'compute size
    'Dim size As Long
    Dim frags As Long
    'size = CLng(header.width) * CLng(header.height)
    frags = CLng(header.bits) / 8
    
    'texture format
    Dim format As GLenum
    Dim intformat As GLint
    Select Case frags
    Case 3
        intformat = 3 'GL_RGB
        format = GL_BGR
    Case 4
        intformat = 4 'GL_RGBA
        format = GL_BGRA
    End Select
        
    'flip texture
    Dim w As Long
    Dim h As Long
    w = header.width
    h = header.height / 6
    
    Dim size As Long
    size = w * h * frags
    
    'allocate buffer
    Dim data() As Byte
    ReDim data(0 To size - 1)
    
    Dim target As GLenum
    target = GL_TEXTURE_CUBE_MAP
    
    'create texture
    Dim handle As GLuint
    glGenTextures 1, handle
    glBindTexture target, handle
    
    'set texture params
    glTexParameteri target, GL_TEXTURE_WRAP_S, GL_CLAMP_TO_EDGE
    glTexParameteri target, GL_TEXTURE_WRAP_T, GL_CLAMP_TO_EDGE
    glTexParameteri target, GL_TEXTURE_MIN_FILTER, GL_LINEAR_MIPMAP_LINEAR
    glTexParameteri target, GL_TEXTURE_MAG_FILTER, GL_LINEAR
    glTexParameterf target, GL_TEXTURE_MAX_ANISOTROPY_EXT, maxaniso
    glTexParameteri target, GL_GENERATE_MIPMAP, GL_TRUE
    
    Dim i As Long
    For i = 0 To 5
        
        'get pixel data
        Get #ff, , data()
        
        Dim uptarget As GLenum
        If i = 5 Then uptarget = GL_TEXTURE_CUBE_MAP_POSITIVE_X
        If i = 4 Then uptarget = GL_TEXTURE_CUBE_MAP_NEGATIVE_X
        If i = 2 Then uptarget = GL_TEXTURE_CUBE_MAP_POSITIVE_Y
        If i = 3 Then uptarget = GL_TEXTURE_CUBE_MAP_NEGATIVE_Y
        If i = 1 Then uptarget = GL_TEXTURE_CUBE_MAP_POSITIVE_Z
        If i = 0 Then uptarget = GL_TEXTURE_CUBE_MAP_NEGATIVE_Z
        
        'upload texture
        glTexImage2D uptarget, 0, intformat, w, h, 0, format, GL_UNSIGNED_BYTE, ByVal VarPtr(data(0))
        
    Next i
    
    'close file
    Close ff
    
    'clean up
    Erase data()
    
    glBindTexture target, 0
    glActiveTexture GL_TEXTURE0
    glClientActiveTexture GL_TEXTURE0
    
    'success
    LoadTGAEnvMap = handle
    Exit Function
errhandler:
    MsgBox err.description, vbCritical, "LoadTGAEnvMap"
End Function
