Attribute VB_Name = "a_Scene"
Option Explicit

'misc
Public mouse_down As Boolean
Public mouse_px As Single
Public mouse_py As Single
Public view_wire As Boolean
Public view_poly As Boolean
Public view_verts As Boolean
Public view_edges As Boolean
Public view_normals As Boolean
Public view_tangents As Boolean
Public view_bounds As Boolean
Public view_bonesys As Boolean
Public view_samples As Boolean
Public view_backfaces As Boolean
Public view_lighting As Boolean
Public view_textures As Boolean
Public view_axis As Boolean
Public view_grids As Boolean
Public view_camanim As Boolean
Private Const axis_size = 1
Private Const grid_step = 50
Private Const grid_size = 1

'camera
Public camcentx As Single
Public camcenty As Single
Public camcentz As Single
Public campanx As Single
Public campany As Single
Public camrotx As Single
Public camroty As Single
Public camzoom As Single
Public camasp As Single
Public Const camnear = 0.01
Public Const camfar = 500

Private drawlock As Boolean


'draws scene
Public Sub DrawScene()
    
    'workaround for race condition on form redraw
    If drawlock = True Then
        'MsgBox "DrawLock violation!", vbExclamation
        Exit Sub
    End If
    drawlock = True
    
    '1p camera
    Dim cam1p As Boolean
    cam1p = view_camanim And bf2ske.loaded And (bf2ske.cambone > -1)
    
    'field of view
    Dim fov As Single
    If cam1p Then
        fov = 60
    Else
        fov = 45
    End If
    
    'clear buffers
    glClearColor bgcolor.r, bgcolor.g, bgcolor.b, bgcolor.a
    glClear GL_COLOR_BUFFER_BIT Or GL_DEPTH_BUFFER_BIT
    
    'setup camera
    glMatrixMode GL_PROJECTION
    glLoadIdentity
    gluPerspective fov, camasp, camnear, camfar
    
    glMatrixMode GL_MODELVIEW
    glLoadIdentity
    If cam1p Then
        Dim m As matrix4
        m = bf2ske.node(bf2ske.cambone).worldmat
        
        gluLookAt -m.m(12), m.m(13), m.m(14), _
                  -m.m(12) + -m.m(8), m.m(13) + m.m(9), m.m(14) + m.m(10), _
                  -m.m(4), m.m(5), m.m(6)
    Else
        glTranslatef campanx, campany, -camzoom
        glRotatef camrotx, 1, 0, 0
        glRotatef camroty, 0, 1, 0
        glTranslatef camcentx, camcenty, camcentz
    End If
    
    GetProjectionInfo
    
    'reset some things
    glEnable GL_DEPTH_TEST
    glEnable GL_CULL_FACE
    glFrontFace GL_CW
    glDisable GL_LIGHTING 'bugfix: after drawok = false
    
    'wireframe
    If view_wire Then
        glPolygonMode GL_FRONT_AND_BACK, GL_LINE
    Else
        glPolygonMode GL_FRONT_AND_BACK, GL_FILL
    End If
    
    'backfaces
    If view_backfaces Then
        glDisable GL_CULL_FACE
    Else
        glEnable GL_CULL_FACE
    End If
    
    'grids
    If view_grids And draw_mode = dm_normal Then
        StartAALine 1.3
        
        'gray gridlines
        glColor3f 0.35, 0.35, 0.35
        Dim i As Long
        For i = -grid_step To grid_step
            If i <> 0 Then
                glBegin GL_LINES
                    glVertex3f -grid_step * grid_size, 0, i * grid_size
                    glVertex3f grid_step * grid_size, 0, i * grid_size
                glEnd
                glBegin GL_LINES
                    glVertex3f i * grid_size, 0, -grid_step * grid_size
                    glVertex3f i * grid_size, 0, grid_step * grid_size
                glEnd
            End If
        Next i
        
        'black gridlines
        glColor3f 0, 0, 0
        glBegin GL_LINES
            glVertex3f -grid_step * grid_size, 0, 0
            glVertex3f grid_step * grid_size, 0, 0
        glEnd
        glBegin GL_LINES
            glVertex3f 0, 0, -grid_step * grid_size
            glVertex3f 0, 0, grid_step * grid_size
        glEnd
        
        EndAALine
    End If
    
    'axis
    If view_axis And draw_mode = dm_normal Then
        StartAALine 2
        DrawPivot axis_size
        EndAALine
    End If
    
    'draw DICE stuff
    glPushMatrix
        glScalef -1, 1, 1
        DrawVisMesh
        DrawColMesh
        DrawSamples
        DrawStdMesh
        DrawTreeMesh
        DrawBF2Skeleton
        DrawConNodes
    glPopMatrix
    DrawObj
    
    'FrostBite
    DrawFbMesh
    
    'draw FHX stuff
    glFrontFace GL_CCW
    DrawFhxGeo
    DrawFhxTri
    DrawFhxRig
    
    'bf2
    DrawOccluder
    
    'done drawing
    glFinish
    
    drawlock = False
End Sub


'zooms in on model
Public Sub ZoomExtends()
    
    campanx = 0
    campany = 0
    
    camcentx = 0
    camcenty = 0
    camcentz = 0
    
    If vmesh.loadok Then
        If selgeom = -1 Then Exit Sub
        If sellod = -1 Then Exit Sub
        With vmesh.geom(selgeom).lod(sellod)
            
            'camrotx = 0
            'camroty = 0
            
            'compute center
            Dim c As float3
            c.X = (.min.X + .max.X) * 0.5
            c.Y = (.min.Y + .max.Y) * 0.5
            c.z = (.min.z + .max.z) * 0.5
            
            'compute radius
            Dim s As Single
            s = Distance(c, .max) * 2
            
            'compute translation
            'Dim v As float3
            'v.x = 0
            'v.y = 0
            'v.z = 1
            
            'Dim r As float3
            'r.x = camrotx
            'r.y = camroty
            
            'Dim p As float3
            'p = Rotate(c, v, r)
            
            'campanx = p.x
            'campany = p.y
            
            camcentx = c.X
            camcenty = -c.Y
            camcentz = -c.z
            camzoom = s
            
            'redraw
            frmMain.picMain_Paint
        End With
    End If
End Sub


'anti-aliased line drawing
Public Sub StartAALine(ByVal width As Single)
    glLineWidth width
    glBlendFunc GL_SRC_ALPHA, GL_ONE_MINUS_SRC_ALPHA
    glEnable GL_LINE_SMOOTH
    glEnable GL_BLEND
    glDepthMask False
End Sub
Public Sub EndAALine()
    glDepthMask True
    glDisable GL_BLEND
    glDisable GL_LINE_SMOOTH
    glLineWidth 1
End Sub


'anti-aliased point drawing
Public Sub StartAAPoint(ByVal size As Single)
    glPointSize size
    glBlendFunc GL_SRC_ALPHA, GL_ONE_MINUS_SRC_ALPHA
    glEnable GL_POINT_SMOOTH
    glEnable GL_BLEND
    glDepthMask False
End Sub
Public Sub EndAAPoint()
    glDepthMask True
    glDisable GL_BLEND
    glDisable GL_POINT_SMOOTH
    glPointSize 1
End Sub


'draws wire box
Public Sub DrawBox(ByRef min As float3, ByRef max As float3)
    
    'bottom square
    glBegin GL_LINE_LOOP
        glVertex3f min.X, min.Y, min.z
        glVertex3f min.X, min.Y, max.z
        glVertex3f max.X, min.Y, max.z
        glVertex3f max.X, min.Y, min.z
    glEnd
    
    'top square
    glBegin GL_LINE_LOOP
        glVertex3f min.X, max.Y, min.z
        glVertex3f min.X, max.Y, max.z
        glVertex3f max.X, max.Y, max.z
        glVertex3f max.X, max.Y, min.z
    glEnd
    
    'vertical lines
    glBegin GL_LINES
        glVertex3f min.X, min.Y, min.z
        glVertex3f min.X, max.Y, min.z
        glVertex3f min.X, min.Y, max.z
        glVertex3f min.X, max.Y, max.z
        
        glVertex3f max.X, min.Y, min.z
        glVertex3f max.X, max.Y, min.z
        glVertex3f max.X, min.Y, max.z
        glVertex3f max.X, max.Y, max.z
    glEnd
End Sub


'draws pivot
Public Sub DrawPivot(ByVal s As Single)
    glBegin GL_LINES
        glColor3f 1, 0, 0
        glVertex3f 0, 0, 0
        glVertex3f s, 0, 0
        
        glColor3f 0, 1, 0
        glVertex3f 0, 0, 0
        glVertex3f 0, s, 0
        
        glColor3f 0, 0, 1
        glVertex3f 0, 0, 0
        glVertex3f 0, 0, s
    glEnd
End Sub


'draws bone node
Public Sub DrawBoneWire(ByVal s As Single)
    
    'x plane
    glBegin GL_LINE_LOOP
        glVertex3f 0, -s, 0
        glVertex3f 0, 0, -s
        glVertex3f 0, s, 0
        glVertex3f 0, 0, s
    glEnd
    
    'y plane
    glBegin GL_LINE_LOOP
        glVertex3f -s, 0, 0
        glVertex3f 0, 0, -s
        glVertex3f s, 0, 0
        glVertex3f 0, 0, s
    glEnd
    
    'z plane
    glBegin GL_LINE_LOOP
        glVertex3f -s, 0, 0
        glVertex3f 0, -s, 0
        glVertex3f s, 0, 0
        glVertex3f 0, s, 0
    glEnd
    
End Sub
