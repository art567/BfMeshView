Attribute VB_Name = "BF2_Bake"
Option Explicit

'bakes texture to lightmap UVs
Public Sub BakeTexture(ByVal geom As Long, ByVal lod As Long, _
                       ByVal w As Long, ByVal h As Long, ByVal uvchan As Long, ByVal padding As Long)
    
    On Error GoTo errhandler
    
    'create pbuffer
    CreatePBuffer w, h
    wglMakeCurrent pbuf_hdc, pbuf_hrc
    wglShareLists frmMain.hglrc, pbuf_hrc
    glTexEnvi GL_TEXTURE_ENV, GL_TEXTURE_ENV_MODE, GL_MODULATE
    
    Dim oldbakemode  As Boolean
    oldbakemode = bakemode
    bakemode = True
    
    Dim oldtexmode As Boolean
    oldtexmode = view_textures
    
    'draw
    glViewport 0, 0, w, h
    
    glClearColor 0.2, 0.2, 0.2, 0
    glClear GL_COLOR_BUFFER_BIT
    glClear GL_DEPTH_BUFFER_BIT
    
    glMatrixMode GL_PROJECTION
    glLoadIdentity
    glOrtho 0, 1, 0, 1, -1, 1
    
    glMatrixMode GL_MODELVIEW
    glLoadIdentity
    
    'bake alpha
    glColorMask False, False, False, True
    view_textures = False
    DrawVisMeshLod vmesh.geom(geom).lod(lod)
    
    'bake color
    glColorMask True, True, True, False
    view_textures = True
    DrawVisMeshLod vmesh.geom(geom).lod(lod)
    
    'allocate readback buffer
    Dim bufferSize As Long
    bufferSize = w * h
    
    Dim buffer() As bgra
    ReDim buffer(0 To bufferSize - 1)
    
    'read back
    glFinish
    glReadPixels 0, 0, w, h, GL_BGRA, GL_UNSIGNED_BYTE, ByVal VarPtr(buffer(0))
    
    'apply padding
    GenPadding w, h, buffer(), padding
    
    'ouput
    Dim fname As String
    fname = App.path & "\" & GetNameFromFileName(vmesh.filename) & "_lod" & sellod & ".tga"
    WriteTGA32 fname, w, h, buffer()
    
    'clean up
    Erase buffer()
    
    'reset
    bakemode = oldbakemode
    view_textures = oldtexmode
    
    glColorMask True, True, True, True
    
    'destroy pbuffer
    DestroyPBuffer
    wglMakeCurrent frmMain.picMain.hDC, frmMain.hglrc
    
    Exit Sub
errhandler:
    MsgBox "BakeTexture" & vbLf & err.description, vbCritical
End Sub

