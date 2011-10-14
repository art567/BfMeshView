Attribute VB_Name = "a_PixelBuffer"
Option Explicit

Public Const WGL_SUPPORT_OPENGL As Long = 8208
Public Const WGL_ACCELERATION As Long = 8195
Public Const WGL_FULL_ACCELERATION As Long = 8231
Public Const WGL_COLOR_BITS As Long = 8212
Public Const WGL_DEPTH_BITS As Long = 8226
Public Const WGL_PBUFFER_WIDTH As Long = 8244
Public Const WGL_PBUFFER_HEIGHT As Long = 8245
Public Const WGL_RED_BITS_ARB As Long = 8213
Public Const WGL_GREEN_BITS_ARB As Long = 8215
Public Const WGL_BLUE_BITS_ARB As Long = 8217
Public Const WGL_ALPHA_BITS As Long = 8219


Public Declare Function glGetPixelFormatAttribiv Lib "glext.dll" (ByVal hDC As Long, ByVal iPixelFormat As Long, _
                                                                  ByVal iLayerPlane As Long, ByVal nAttributes As Long, _
                                                                  ByRef piAttributes As Long, ByRef piValues As Long) As Boolean

Public Declare Function glGetPixelFormatAttribfv Lib "glext.dll" (ByVal hDC As Long, ByVal iPixelFormat As Long, _
                                                                  ByVal iLayerPlane As Long, nAttributes As Long, _
                                                                  ByRef piAttributes, ByRef pfValues As Single) As Boolean

Public Declare Function glChoosePixelFormat Lib "glext.dll" (ByVal hDC As Long, ByVal piAttribIList As Long, _
                                                             ByVal pfAttribFList As Long, ByVal nMaxFormats As Long, _
                                                             ByVal piFormats As Long, ByVal nNumFormats As Long) As Boolean

Public Declare Function glCreatePbuffer Lib "glext.dll" (ByVal hDC As Long, ByVal iPixelFormat As Long, _
                                                         ByVal iWidth As Long, ByVal iHeight As Long, _
                                                         ByVal piAttribList As Long) As Long

Public Declare Function glGetPbufferDC Lib "glext.dll" (ByVal hPbuffer As Long) As Long

Public Declare Function glReleasePbufferDC Lib "glext.dll" (ByVal hPbuffer As Long, ByVal hDC As Long) As Long

Public Declare Function glDestroyPbuffer Lib "glext.dll" (ByVal hPbuffer As Long) As Boolean

Public Declare Function glQueryPbuffer Lib "glext.dll" (ByVal hPbuffer As Long, ByVal iAttribute As Long, ByRef piValue As Long) As Boolean


Public pbuf_hrc As Long     'rendering context handle
Public pbuf_hdc As Long     'device context handle
Private pbuf_pbo As Long    'pixel buffer object handle
Private pbuf_w As Long
Private pbuf_h As Long


'creates pixel buffer
Public Sub CreatePBuffer(ByVal w As Long, ByVal h As Long)
    pbuf_w = w
    pbuf_h = h
    
    'get current device context
    Dim currentDC As Long
    currentDC = wglGetCurrentDC()
    If currentDC = 0 Then
        MsgBox "wglGetCurrentDC failed", vbCritical
        Exit Sub
    End If
    
    'choose pixel format
    Dim pixelFormat As Long
    Dim numFormats As Long
    
    Dim attribf(0 To 1) As Single
    attribf(0) = 0
    attribf(1) = 0
    
    Dim attribi(0 To 9) As Long
    attribi(0) = WGL_SUPPORT_OPENGL
    attribi(1) = GL_TRUE
    
    attribi(2) = WGL_ACCELERATION
    attribi(3) = WGL_FULL_ACCELERATION
    
    attribi(4) = WGL_COLOR_BITS
    attribi(5) = 32
    
    attribi(6) = WGL_ALPHA_BITS
    attribi(7) = 8
    
    attribi(8) = WGL_DEPTH_BITS
    attribi(9) = 16
    
    attribi(10) = 0
    attribi(11) = 0
    
    If glChoosePixelFormat(currentDC, VarPtr(attribi(0)), VarPtr(attribf(0)), 1, VarPtr(pixelFormat), VarPtr(numFormats)) = 0 Then
        MsgBox "wglChoosePixelFormat failed", vbCritical
        Exit Sub
    End If
    
    'create pixel buffer object
    pbuf_pbo = glCreatePbuffer(currentDC, pixelFormat, pbuf_w, pbuf_h, 0)
    If pbuf_pbo = 0 Then
        MsgBox "wglCreatePbuffer failed", vbCritical
        Exit Sub
    End If
    
    'get device context
    pbuf_hdc = glGetPbufferDC(pbuf_pbo)
    If pbuf_hdc = 0 Then
        MsgBox "wglGetPbufferDC failed", vbCritical
        Exit Sub
    End If
    
    'create rendering context
    pbuf_hrc = wglCreateContext(pbuf_hdc)
    If pbuf_hrc = 0 Then
        MsgBox "wglCreateContext failed", vbCritical
        Exit Sub
    End If
    
    'set pbuffer size
    glQueryPbuffer pbuf_pbo, WGL_PBUFFER_WIDTH, pbuf_w
    glQueryPbuffer pbuf_pbo, WGL_PBUFFER_HEIGHT, pbuf_h
    
End Sub


'destroys pixel buffer
Public Sub DestroyPBuffer()
    
    'ensure we are not current
    wglMakeCurrent 0, 0
    
    'release rendering context
    If pbuf_hrc Then
        If wglDeleteContext(pbuf_hrc) = 0 Then
            MsgBox "wglDeleteContext failed", vbCritical
        End If
        pbuf_hrc = 0
    End If
    
    'release device context
    If pbuf_hdc Then
        If glReleasePbufferDC(pbuf_pbo, pbuf_hdc) = 0 Then
            MsgBox "wglReleasePbufferDC failed", vbCritical
        End If
        pbuf_hdc = 0
    End If
    
    'destroy pixel buffer object
    If pbuf_pbo Then
        If Not glDestroyPbuffer(pbuf_pbo) Then
            MsgBox "wglDestroyPbuffer failed", vbCritical
        End If
        pbuf_pbo = 0
    End If
End Sub

