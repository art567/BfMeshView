Attribute VB_Name = "a_GLSL"
Option Explicit


Public Declare Function glCreateShader Lib "glext.dll" (ByVal shaderType As GLenum) As GLuint
Public Declare Function glCreateProgram Lib "glext.dll" () As GLuint
Public Declare Sub glShaderSource Lib "glext.dll" (ByVal shader As GLuint, ByVal count As GLsizei, _
                                                   sPtr As Long, ByVal lenPtr As Long)
Public Declare Sub glCompileShader Lib "glext.dll" (ByVal shader As GLuint)
Public Declare Sub glGetObjectParameteriv Lib "glext.dll" (ByVal shader As GLuint, ByVal param As GLenum, out As GLint)
Public Declare Sub glAttachShader Lib "glext.dll" (ByVal program As GLuint, ByVal shader As GLuint)
Public Declare Sub glLinkProgram Lib "glext.dll" (ByVal program As GLuint)
Public Declare Sub glUseProgram Lib "glext.dll" (ByVal program As GLuint)
Public Declare Sub glDeleteProgram Lib "glext.dll" (ByVal program As GLuint)
Public Declare Sub glDeleteShader Lib "glext.dll" (ByVal shader As GLuint)
Public Declare Sub glGetShaderInfoLog Lib "glext.dll" (ByVal shader As GLuint, ByVal maxLength As GLsizei, _
                                                       Length As GLsizei, ByVal sPtr As Long)

Public Declare Function glGetUniformLocation Lib "glext.dll" (ByVal program As GLuint, ByVal name As String) As GLuint
Public Declare Sub glUniform1i Lib "glext.dll" (ByVal location As GLint, ByVal v As GLint)
Public Declare Sub glUniform1f Lib "glext.dll" (ByVal location As GLint, ByVal v As GLfloat)
Public Declare Sub glUniform3f Lib "glext.dll" (ByVal location As GLint, _
                                                ByVal v0 As GLfloat, ByVal v1 As GLfloat, ByVal v2 As GLfloat)

Public Declare Sub glUniformMatrix4fv Lib "glext.dll" (ByVal location As GLint, ByVal count As GLsizei, _
                                                       ByVal transpose As GLboolean, value As GLfloat)

Public Type shader
    vert As GLuint
    frag As GLuint
    prog As GLuint
End Type


Public Sub SetNodeTransforms(ByRef sh As shader, ByRef name As String)
    Dim loc As GLuint
    loc = glGetUniformLocation(sh.prog, name)
    If loc <> -1 Then
        glUniformMatrix4fv loc, nodetransformnum, GL_FALSE, nodetransform(0).m(0)
    End If
End Sub

'converts string to char array
Private Function StrToChar(ByRef str As String) As Long
    Static buff(0 To 4096 - 1) As Byte
    
    Dim slen As Long
    slen = Len(str)
    
    Dim i As Long
    For i = 0 To slen - 1
        buff(i) = AscB(Mid(str, i + 1, 1))
    Next i
    buff(slen) = 0
    
    StrToChar = VarPtr(buff(0))
End Function


'creates shader program
Public Function CreateProgram(ByRef sh As shader, ByRef vert As String, ByRef frag As String) As Boolean
    
    If sh.prog Then
        DeleteProgram sh
    End If
    
    Dim vertlen As Long
    Dim fraglen As Long
    vertlen = Len(vert) + 1
    fraglen = Len(frag) + 1
    
    Dim vertshader As GLuint
    Dim fragshader As GLuint
    Dim program As GLuint
    Dim err As GLint
    
    'create shaders
    vertshader = glCreateShader(GL_VERTEX_SHADER)
    fragshader = glCreateShader(GL_FRAGMENT_SHADER)
    
    'pass shader source
    glShaderSource vertshader, 1, StrToChar(vert), VarPtr(vertlen)
    glShaderSource fragshader, 1, StrToChar(frag), VarPtr(fraglen)
    
    'compile vert shader code
    glCompileShader vertshader
    glGetObjectParameteriv vertshader, GL_COMPILE_STATUS, err
    If err = 0 Then
        MsgBox "Vertex shader compile error:" & vbLf & GetErrorLog(vertshader)
        Exit Function
    End If
    
    'compile frag shader code
    glCompileShader fragshader
    glGetObjectParameteriv fragshader, GL_COMPILE_STATUS, err
    If err = 0 Then
        MsgBox "Fragment shader compile error:" & vbLf & GetErrorLog(fragshader)
        Exit Function
    End If
    
    'create program
    program = glCreateProgram
    
    'attach shaders to program
    glAttachShader program, vertshader
    glAttachShader program, fragshader
    
    'link program
    glLinkProgram program
    glGetObjectParameteriv program, GL_LINK_STATUS, err
    If err = 0 Then
        MsgBox "Shader program link error:" & vbLf & GetErrorLog(program)
        Exit Function
    End If
    
    'bind program
    glUseProgram program
    
    'texture handles
    Dim i As Long
    For i = 0 To 7
        Dim loc As GLuint
        loc = glGetUniformLocation(program, "texture" & i)
        If loc <> -1 Then glUniform1i loc, i
    Next i
    
    'clean up
    glUseProgram 0
    
    sh.vert = vertshader
    sh.frag = fragshader
    sh.prog = program
    
    CreateProgram = True
End Function


'set uniform
Public Sub SetUniform3f(ByRef sh As shader, ByRef name As String, ByRef val As float3)
    Dim loc As GLuint
    loc = glGetUniformLocation(sh.prog, name)
    If loc <> -1 Then glUniform3f loc, val.X, val.y, val.z
End Sub

'set uniform
Public Sub SetUniform1f(ByRef sh As shader, ByRef name As String, ByVal val As Single)
    Dim loc As GLuint
    loc = glGetUniformLocation(sh.prog, name)
    If loc <> -1 Then glUniform1f loc, val
End Sub


'deletes program
Public Function DeleteProgram(ByRef sh As shader)
    glDeleteProgram sh.prog
    glDeleteShader sh.vert
    glDeleteShader sh.frag
    sh.prog = 0
    sh.vert = 0
    sh.frag = 0
End Function



'prints shader/program error log
Private Function GetErrorLog(ByVal obj As GLuint) As String
    
    Dim slen As GLint
    glGetObjectParameteriv obj, GL_INFO_LOG_LENGTH, slen
    
    If slen < 1 Then
        GetErrorLog = "No error string???"
        Exit Function
    End If
    
    Dim buffer() As Byte
    ReDim buffer(0 To slen - 1)
    
    Dim chars As GLsizei
    glGetShaderInfoLog obj, slen, chars, VarPtr(buffer(0))
    
    Dim str As String
    Dim i As Long
    For i = 0 To chars - 1
        str = str & Chr(buffer(i))
    Next i
    
    GetErrorLog = str
End Function



