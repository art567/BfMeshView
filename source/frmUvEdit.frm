VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUvEdit 
   Caption         =   "UV Editor"
   ClientHeight    =   6435
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8325
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmUvEdit.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   429
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   555
   Begin MSComctlLib.ImageList imlTools 
      Left            =   6600
      Top             =   3000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUvEdit.frx":1042
            Key             =   "select"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUvEdit.frx":11AC
            Key             =   "move"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUvEdit.frx":1306
            Key             =   "scale"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUvEdit.frx":146C
            Key             =   "uvxneg"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUvEdit.frx":17BE
            Key             =   "uvxnpos"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUvEdit.frx":1B10
            Key             =   "uvyneg"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUvEdit.frx":1E62
            Key             =   "uvypos"
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraTools 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      Caption         =   "Tools"
      Height          =   780
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   5655
      Begin VB.ComboBox cbbMaterial 
         Height          =   315
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   60
         Width           =   1935
      End
      Begin VB.CommandButton cmdScale 
         Caption         =   "Scale"
         Height          =   315
         Left            =   3720
         TabIndex        =   6
         Top             =   420
         Width           =   855
      End
      Begin VB.CommandButton cmdRotate 
         Caption         =   "Rotate"
         Height          =   315
         Left            =   2760
         TabIndex        =   5
         Top             =   420
         Width           =   855
      End
      Begin VB.ComboBox cbbChannel 
         Height          =   315
         Left            =   60
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   60
         Width           =   1935
      End
      Begin VB.CommandButton cmdCollapse 
         Caption         =   "Collapse"
         Height          =   315
         Left            =   4680
         TabIndex        =   2
         Top             =   420
         Width           =   855
      End
      Begin MSComctlLib.Toolbar tlbTools 
         Height          =   330
         Left            =   60
         TabIndex        =   4
         Top             =   405
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "imlTools"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   8
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "select"
               Object.ToolTipText     =   "Select UVs"
               ImageIndex      =   1
               Style           =   2
               Value           =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "move"
               Object.ToolTipText     =   "Move UVs"
               ImageIndex      =   2
               Style           =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "scale"
               Object.ToolTipText     =   "Scale UVs"
               ImageIndex      =   3
               Style           =   2
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "uvxneg"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "uvxpos"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "uvyneg"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "uvypos"
               ImageIndex      =   7
            EndProperty
         EndProperty
      End
   End
   Begin VB.PictureBox picMain 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   4500
      Left            =   840
      ScaleHeight     =   300
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   300
      TabIndex        =   0
      Top             =   1080
      Width           =   4500
   End
End
Attribute VB_Name = "frmUvEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum e_toolmode
    tool_select = 0
    tool_move = 1
    tool_scale = 2
End Enum

Private zoom As Single
Private offx As Single
Private offy As Single
Private mousedown As Boolean
Private mousex As Long
Private mousey As Long
Private dragx As Single 'mouse coordinates at start of drag
Private dragy As Single 'mouse coordinates at start of drag
Private dsx As Single 'zoom scale
Private dsy As Single 'zoom scale
Private toolmode As e_toolmode
Private uvchan As Long
Private uvmatsel As Long

Private sel_vis As Boolean
Private sel_minx As Single
Private sel_miny As Single
Private sel_maxx As Single
Private sel_maxy As Single
                
Private keyctrl As Boolean
Private keyalt As Boolean

Private hglrc As Long

Private Sub Form_Load()
    
    Me.fraTools.BackColor = &H8000000F
    
    'setup opengl
    Dim pfd As PIXELFORMATDESCRIPTOR
    Dim fmt As Long
    pfd.nSize = Len(pfd)
    pfd.nVersion = 1
    pfd.dwFlags = PFD_SUPPORT_OPENGL Or PFD_DRAW_TO_WINDOW Or PFD_DOUBLEBUFFER Or PFD_TYPE_RGBA
    pfd.iPixelType = PFD_TYPE_RGBA
    pfd.cColorBits = 24
    pfd.cDepthBits = 0
    pfd.iLayerType = PFD_MAIN_PLANE
    fmt = ChoosePixelFormat(Me.picMain.hdc, pfd)
    If fmt = 0 Then
        MsgBox "OpenGL initalization failed.", vbCritical
        Exit Sub
    End If
    fmt = SetPixelFormat(Me.picMain.hdc, fmt, pfd)
    hglrc = wglCreateContext(Me.picMain.hdc)
    wglShareLists frmMain.hglrc, hglrc
    wglMakeCurrent Me.picMain.hdc, hglrc
    
    'default states
    glTexEnvi GL_TEXTURE_ENV, GL_TEXTURE_ENV_MODE, GL_MODULATE
    
    uveditor_isloaded = True
    
    offx = 50
    offy = 50
    zoom = 1
    uvchan = 0
    toolmode = 0
    
    FillChannelList
    FillMaterialList
    
    SetTopMostWindow Me.hWnd, True
    
    Center Me
End Sub

Private Sub Form_Resize()
    If Not Me.WindowState = vbMinimized Then
        If Me.width < 200 * 15 Then Me.width = 200 * 15
        If Me.height < 100 * 15 Then Me.height = 100 * 15
        Me.picMain.Move 3, Me.fraTools.height, Me.ScaleWidth - 6, Me.ScaleHeight - fraTools.height - 3
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not app_exit Then
        Cancel = True
        Me.Hide
        Exit Sub
    End If
    If hglrc Then
        wglMakeCurrent 0, 0
        wglDeleteContext hglrc
        hglrc = 0
    End If
End Sub

Private Sub picMain_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyControl Then keyctrl = True
    If KeyCode = vbKeyMenu Then keyalt = True
End Sub

Private Sub picMain_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyControl Then keyctrl = False
    If KeyCode = vbKeyMenu Then keyalt = False
End Sub

Private Sub picMain_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mousedown = True
    mousex = X
    mousey = Y
    
    dragx = X
    dragy = Y
    
    picMain_Paint
End Sub

Private Sub picMain_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mousedown Then
        
        Dim vx As Single
        Dim vy As Single
        vx = (X - mousex)
        vy = (Y - mousey)
        
        If Button = vbLeftButton Then
            Select Case toolmode
            Case tool_select
                
                sel_vis = True
                sel_minx = (min(dragx, X) - offx) / dsx
                sel_miny = (min(dragy, Y) - offy) / dsy
                sel_maxx = (max(dragx, X) - offx) / dsx
                sel_maxy = (max(dragy, Y) - offy) / dsy
                
            Case tool_move
                MoveVerts vx / dsx, vy / dsy
                
            Case tool_scale
                ScaleVerts vx / dsx, vy / dsy
                
            End Select
        End If
        
        If Button = vbRightButton Then
            
            Dim cx As Single
            Dim cy As Single
            cx = TFXi(dragx)
            cy = TFYi(dragy)
            
            zoom = zoom - (vy * 0.01 * zoom)
            If zoom < 0.01 Then zoom = 0.01
            If zoom > 100 Then zoom = 100
            
            Dim ncx As Single
            Dim ncy As Single
            ncx = TFXi(dragx)
            ncy = TFYi(dragy)
            
            offx = offx + (ncx - cx) * zoom
            offy = offy + (ncy - cy) * zoom
            
        End If
        
        If Button = vbMiddleButton Then
            offx = offx + vx
            offy = offy + vy
        End If
        
        picMain_Paint
        frmMain.picMain_Paint
    End If
    mousex = X
    mousey = Y
End Sub


Private Sub picMain_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    'ignore MouseUp after minimizing/maximizing window
    If Not mousedown Then Exit Sub
    
    'selection
    If Button = vbLeftButton Then
        If toolmode = tool_select Then
            
            If sel_vis Then
                SelVerts sel_minx, sel_miny, sel_maxx, sel_maxy
            Else
                SelVerts ((X - 3) - offx) / dsx, ((Y - 3) - offy) / dsy, _
                         ((X + 3) - offx) / dsx, ((Y + 3) - offy) / dsy
            End If
            
        End If
    End If
    
    sel_vis = False
    mousedown = False
    picMain_Paint
    frmMain.picMain_Paint
End Sub

Private Sub cbbMaterial_Click()
    uvmatsel = Me.cbbMaterial.ListIndex
    
    ClearVertSelection
    picMain_Paint
    frmMain.picMain_Paint
End Sub

Private Sub cbbChannel_Click()
    uvchan = Me.cbbChannel.ListIndex
    If uvchan < 0 Then uvchan = 0
    If uvchan > 4 Then uvchan = 4
    If (uvchan < 4) Then
        Me.cbbMaterial.BackColor = &H80000005
        Me.cbbMaterial.Enabled = True
        uvmatsel = Me.cbbMaterial.ListIndex
    Else
        Me.cbbMaterial.BackColor = &H8000000F
        Me.cbbMaterial.Enabled = False
        uvmatsel = -1
    End If
    
    SetVertFlags
    ClearVertSelection
    picMain_Paint
    frmMain.picMain_Paint
End Sub

'fills dropdown lists
Public Sub FillChannelList()
    With vmesh
        If Not .loadok Then Exit Sub
        Dim i As Long
        Me.cbbChannel.Clear
        For i = 0 To .vertattribnum - 1
            If Not .vertattrib(i).flag = 255 Then
                Select Case .vertattrib(i).usage
                Case 5: Me.cbbChannel.AddItem "UV 1 (Base)"
                Case 261: Me.cbbChannel.AddItem "UV 2 (Detail)"
                Case 517: Me.cbbChannel.AddItem "UV 3 (Dirt)"
                Case 773: Me.cbbChannel.AddItem "UV 4 (Crack)"
                Case 1029: Me.cbbChannel.AddItem "UV 5 (Lightmap)"
                End Select
            End If
        Next i
        Me.cbbChannel.ListIndex = uvchan
        
        SetVertFlags
        ClearVertSelection
    End With
End Sub

'fills dropdown lists
Public Sub FillMaterialList()
    With vmesh
        If Not .loadok Then Exit Sub
        
        Me.cbbMaterial.Clear
        With .geom(selgeom).lod(sellod)
            Dim i As Long
            For i = 0 To .matnum - 1
                Me.cbbMaterial.AddItem "Material " & i
            Next i
        End With
        Me.cbbMaterial.ListIndex = 0
        'uvmatsel = -1
        
        'SetVertFlags
        'ClearVertSelection
    End With
End Sub

'selects UV vertices within rectangular boundaries
Private Sub SelVerts(ByVal minx As Single, ByVal miny As Single, ByVal maxx As Single, ByVal maxy As Single)
Dim i As Long
    
    'scale uv space
    With vmesh
        If Not .loadok Then Exit Sub
        SetVertFlags
        
        Dim stride As Long
        Dim uvoffset As Long
        
        uvoffset = BF2MeshGetTexcOffset(uvchan)
        stride = .vertstride / 4
        
        For i = 0 To .vertnum - 1
            If .vertflag(i) Then
            
                Dim X As Single
                Dim Y As Single
                X = .vert(i * stride + uvoffset + 0)
                Y = .vert(i * stride + uvoffset + 1)
                
                'clear vert selection
                If Not keyctrl And Not keyalt Then
                    .vertsel(i) = 0
                End If
                If X >= minx Then
                    If X <= maxx Then
                        If Y > miny Then
                            If Y < maxy Then
                                
                                If keyalt Then
                                    .vertsel(i) = 0
                                Else
                                    .vertsel(i) = 1
                                End If
                                
                            End If
                        End If
                    End If
                End If
                
            End If
        Next i
        
    End With
    
End Sub


'moves selected UV vertices
Private Sub MoveVerts(ByVal vx As Single, ByVal vy As Single)
Dim i As Long
    With vmesh
        If Not .loadok Then Exit Sub
        SetVertFlags
        
        Dim stride As Long
        Dim uvoffset As Long
        
        uvoffset = BF2MeshGetTexcOffset(uvchan)
        stride = .vertstride / 4
        
        For i = 0 To .vertnum - 1
            If .vertflag(i) Then
            If .vertsel(i) Then
                
                Dim X As Single
                Dim Y As Single
                
                X = .vert(i * stride + uvoffset + 0)
                Y = .vert(i * stride + uvoffset + 1)
                
                X = X + vx
                Y = Y + vy
                
                .vert(i * stride + uvoffset + 0) = X
                .vert(i * stride + uvoffset + 1) = Y
                
            End If
            End If
        Next i
    End With
End Sub


'scales selected UV vertices
Private Sub ScaleVerts(ByVal vx As Single, ByVal vy As Single)
Dim i As Long
    With vmesh
        If Not .loadok Then Exit Sub
        SetVertFlags
        
        Dim sx As Single
        Dim sy As Single
        sx = 1 + vx
        sy = 1 + vy
        
        Dim stride As Long
        Dim uvoffset As Long
        
        uvoffset = BF2MeshGetTexcOffset(uvchan)
        stride = .vertstride / 4
        
        For i = 0 To .vertnum - 1
            If .vertflag(i) Then
            If .vertsel(i) Then
                
                .vert(i * stride + uvoffset + 0) = .vert(i * stride + uvoffset + 0) * sx
                .vert(i * stride + uvoffset + 1) = .vert(i * stride + uvoffset + 1) * sy
                
            End If
            End If
        Next i
    End With
End Sub


'rotates UVs 90 degrees
Private Sub cmdRotate_Click()
Dim i As Long
    With vmesh
        If Not .loadok Then Exit Sub
        SetVertFlags
        
        Dim stride As Long
        Dim uvoffset As Long
        uvoffset = BF2MeshGetTexcOffset(uvchan)
        stride = .vertstride / 4
        
        'collapse vertices
        For i = 0 To .vertnum - 1
            If .vertflag(i) Then
            If .vertsel(i) Then
                
                Dim X As Single
                Dim Y As Single
                
                X = .vert(i * stride + uvoffset + 0)
                Y = .vert(i * stride + uvoffset + 1)
                
                .vert(i * stride + uvoffset + 0) = Y
                .vert(i * stride + uvoffset + 1) = X
                
            End If
            End If
        Next i
        
    End With
    
    picMain_Paint
    frmMain.picMain_Paint
    
    Me.picMain.SetFocus
End Sub


'scale
Private Sub cmdScale_Click()
    
    Dim str As String
    str = InputBox("Scale Factor:", "Scale", 1)
    If Len(str) = 0 Then Exit Sub
    
    Dim s As Single
    s = val(str)
    
    With vmesh
        If Not .loadok Then Exit Sub
        SetVertFlags
        
        Dim stride As Long
        Dim uvoffset As Long
        uvoffset = BF2MeshGetTexcOffset(uvchan)
        stride = .vertstride / 4
        
        Dim i As Long
        For i = 0 To .vertnum - 1
            If .vertflag(i) Then
            If .vertsel(i) Then
                
                .vert(i * stride + uvoffset + 0) = .vert(i * stride + uvoffset + 0) * s
                .vert(i * stride + uvoffset + 1) = .vert(i * stride + uvoffset + 1) * s
                
            End If
            End If
        Next i
        
    End With
    
    picMain_Paint
    frmMain.picMain_Paint
    
    Me.picMain.SetFocus
End Sub


'collapses selected UV vertices
Private Sub cmdCollapse_Click()
Dim i As Long
    With vmesh
        If Not .loadok Then Exit Sub
        SetVertFlags
        
        Dim stride As Long
        Dim uvoffset As Long
        uvoffset = BF2MeshGetTexcOffset(uvchan)
        stride = .vertstride / 4
        
        'determine selected vertex bounds
        Dim minx As Single
        Dim miny As Single
        Dim maxx As Single
        Dim maxy As Single
        minx = 9999
        miny = 9999
        maxx = -9999
        maxy = -9999
        For i = 0 To .vertnum - 1
            If .vertsel(i) Then
                
                Dim X As Single
                Dim Y As Single
                
                X = .vert(i * stride + uvoffset + 0)
                Y = .vert(i * stride + uvoffset + 1)
                
                minx = min(minx, X)
                miny = min(miny, Y)
                maxx = max(maxx, X)
                maxy = max(maxy, Y)
            End If
        Next i
        
        'compute selection center
        Dim cx As Single
        Dim cy As Single
        cx = (minx + maxx) * 0.5
        cy = (miny + maxy) * 0.5
        
        'collapse vertices
        For i = 0 To .vertnum - 1
            If .vertflag(i) Then
            If .vertsel(i) Then
                .vert(i * stride + uvoffset + 0) = cx
                .vert(i * stride + uvoffset + 1) = cy
            End If
            End If
        Next i
        
    End With
    
    picMain_Paint
    frmMain.picMain_Paint
    
    Me.picMain.SetFocus
End Sub


'clears vertex selection
Private Sub ClearVertSelection()
    With vmesh
        If Not .loadok Then Exit Sub
        Dim i As Long
        For i = 0 To .vertnum - 1
            .vertsel(i) = 0
        Next i
    End With
End Sub


'sets the vertex flags of the currently selected geom+lod+mat
Public Sub SetVertFlags()
    
    On Error GoTo errhandler
    
    Dim i As Long
    With vmesh
        If Not .loadok Then Exit Sub
        If selgeom < 0 Then Exit Sub
        If sellod < 0 Then Exit Sub
        
        'clear vert flags
        For i = 0 To .vertnum - 1
            .vertflag(i) = 0
        Next i
        
        'get some info
        Dim stride As Long
        Dim uvoffset As Long
        uvoffset = BF2MeshGetTexcOffset(uvchan)
        stride = .vertstride / 4
        
        With .geom(selgeom).lod(sellod)
            
            Dim matmin As Long
            Dim matmax As Long
            If uvmatsel < 0 Then
                matmin = 0
                matmax = .matnum - 1
            Else
                matmin = uvmatsel
                matmax = uvmatsel
            End If
            
            Dim m As Long
            For m = matmin To matmax
                With .mat(m)
                    Dim facenum As Long
                    facenum = .inum / 3
                    
                    For i = 0 To facenum - 1
                        
                        Dim v1 As Long
                        Dim v2 As Long
                        Dim v3 As Long
                        v1 = .vstart + vmesh.Index(.istart + (i * 3) + 0)
                        v2 = .vstart + vmesh.Index(.istart + (i * 3) + 1)
                        v3 = .vstart + vmesh.Index(.istart + (i * 3) + 2)
                        
                        Dim f1 As Long
                        Dim f2 As Long
                        Dim f3 As Long
                        f1 = (v1 * stride) + uvoffset
                        f2 = (v2 * stride) + uvoffset
                        f3 = (v3 * stride) + uvoffset
                        
                        vmesh.vertflag(v1) = 1
                        vmesh.vertflag(v2) = 1
                        vmesh.vertflag(v3) = 1
                    Next i
                End With
            Next m
        End With
    End With
    
    Exit Sub
errhandler:
    Me.Caption = "SetVertFlags Error: " & err.description
    On Error GoTo 0
End Sub


'redraws UV view
Public Sub picMain_Paint()
    If frmMain.blockdraw Then Exit Sub
    If Not Me.Visible Then Exit Sub
    
    On Error GoTo errhandler:
    
    'update draw scale
    dsx = zoom * 200
    dsy = zoom * 200
    
    DrawGL
    SwapBuffers Me.picMain.hdc
    
    Exit Sub
errhandler:
    Me.Caption = "picMain_Paint Error: " & err.description
    On Error GoTo 0
End Sub

'draws view with OpenGL
Private Sub DrawGL()
    If hglrc = 0 Then Exit Sub
    Dim w As Long
    Dim h As Long
    w = Me.picMain.ScaleWidth
    h = Me.picMain.ScaleHeight
    If w = 0 Then Exit Sub
    If h = 0 Then Exit Sub
    
    wglMakeCurrent Me.picMain.hdc, hglrc
    
    glViewport 0, 0, w, h
    glClearColor 0.25, 0.25, 0.25, 0
    glClear GL_COLOR_BUFFER_BIT
    
    glMatrixMode GL_PROJECTION
    glLoadIdentity
    glOrtho 0, w, h, 0, -1, 1
    
    glMatrixMode GL_MODELVIEW
    glLoadIdentity
    glTranslatef offx, offy, 0
    'glScalef zoom, zoom, 1
    glScalef dsx, dsy, 1
    
    glDisable GL_LIGHTING
    glDisable GL_CULL_FACE
    
    Dim i As Long
    With vmesh
        If Not .loadok Then Exit Sub
        If selgeom < 0 Then Exit Sub
        If sellod < 0 Then Exit Sub
        
        Dim stride As Long
        Dim uvoffset As Long
        uvoffset = BF2MeshGetTexcOffset(uvchan)
        stride = .vertstride / 4
        
        'draw triangles
        Dim s As Long
        Dim v As Long
        With .geom(selgeom).lod(sellod)
            If uvchan = 4 Then
                
                'draw background
                glColor3f 0.3, 0.3, 0.3
                glBegin GL_QUADS
                    glVertex2f 0, 0
                    glVertex2f 0, 1
                    glVertex2f 1, 1
                    glVertex2f 1, 0
                glEnd
                
                glEnableClientState GL_VERTEX_ARRAY
                Dim m As Long
                For m = 0 To .matnum - 1
                    With .mat(m)
                        glColor3f 0.5, 0.5, 0.5
                        
                        s = .istart
                        v = (.vstart * stride) + uvoffset
                        
                        'draw solid
                        glVertexPointer 2, GL_FLOAT, vmesh.vertstride, ByVal VarPtr(vmesh.vert(v))
                        glDrawElements GL_TRIANGLES, .inum, GL_UNSIGNED_SHORT, ByVal VarPtr(vmesh.Index(s))
                        
                        'draw wire
                        StartAALine 1.333: glBlendFunc GL_ALPHA, GL_ONE_MINUS_SRC_COLOR
                        glPolygonMode GL_FRONT_AND_BACK, GL_LINE
                        glColor3f 0.5, 0.5, 0.5
                        glDisable GL_TEXTURE_2D
                        glColor4f 1, 1, 1, 0.1
                        glDrawElements GL_TRIANGLES, .inum, GL_UNSIGNED_SHORT, ByVal VarPtr(vmesh.Index(s))
                        glPolygonMode GL_FRONT_AND_BACK, GL_FILL
                        EndAALine
                    End With
                Next m
                glDisableClientState GL_VERTEX_ARRAY
            Else
                With .mat(uvmatsel)
                
                    Dim ch As Long
                    ch = uvchan
                    If ch > .mapnum - 1 Then ch = .mapnum - 1
                    If ch > 4 Then ch = 4
                    If .texmapid(ch) Then
                        BindTexture .texmapid(ch)
                    Else
                        UnbindTexture
                    End If
                    
                    'draw background
                    glColor3f 0.5, 0.5, 0.5
                    glBegin GL_QUADS
                        glTexCoord2f 0, 0: glVertex2f 0, 0
                        glTexCoord2f 0, 1: glVertex2f 0, 1
                        glTexCoord2f 1, 1: glVertex2f 1, 1
                        glTexCoord2f 1, 0: glVertex2f 1, 0
                    glEnd
                    
                    s = .istart
                    v = (.vstart * stride) + uvoffset
                    
                    glVertexPointer 2, GL_FLOAT, vmesh.vertstride, ByVal VarPtr(vmesh.vert(v))
                    glTexCoordPointer 2, GL_FLOAT, vmesh.vertstride, ByVal VarPtr(vmesh.vert(v))
                    
                    glEnableClientState GL_VERTEX_ARRAY
                    glEnableClientState GL_TEXTURE_COORD_ARRAY
                    
                    'draw solid
                    glColor3f 1, 1, 1
                    glDrawElements GL_TRIANGLES, .inum, GL_UNSIGNED_SHORT, ByVal VarPtr(vmesh.Index(s))
                    
                    'draw wire
                    StartAALine 1.333: glBlendFunc GL_ALPHA, GL_ONE_MINUS_SRC_COLOR
                    glPolygonMode GL_FRONT_AND_BACK, GL_LINE
                    glDisable GL_TEXTURE_2D
                    glColor4f 1, 1, 1, 0.1
                    glDrawElements GL_TRIANGLES, .inum, GL_UNSIGNED_SHORT, ByVal VarPtr(vmesh.Index(s))
                    glPolygonMode GL_FRONT_AND_BACK, GL_FILL
                    EndAALine
                    
                    glDisableClientState GL_VERTEX_ARRAY
                    glDisableClientState GL_TEXTURE_COORD_ARRAY
                End With
            End If
                        
        End With
        
        'draw vertices
        SetVertFlags
        StartAAPoint 4: glBlendFunc GL_ALPHA, GL_ONE_MINUS_SRC_COLOR
        glBegin GL_POINTS
        For i = 0 To .vertnum - 1
            If .vertflag(i) Then
                If .vertsel(i) Then
                    glColor3f 1, 0, 0
                Else
                    glColor3f 1, 1, 1
                End If
                glVertex2fv .vert(i * stride + uvoffset)
            End If
        Next i
        glEnd
        EndAAPoint
        
        'draw selection rectangle
        If sel_vis Then
            glBlendFunc GL_SRC_ALPHA, GL_ONE_MINUS_SRC_ALPHA
            glEnable GL_BLEND
            glColor4f 1, 1, 1, 0.1
            glBegin GL_QUADS
                glVertex2f sel_minx, sel_miny
                glVertex2f sel_minx, sel_maxy
                glVertex2f sel_maxx, sel_maxy
                glVertex2f sel_maxx, sel_miny
            glEnd
            glDisable GL_BLEND
        End If
        
    End With
    
End Sub


Private Function TFX(ByVal X As Single) As Single
    TFX = (X * zoom) + offx
End Function

Private Function TFY(ByVal Y As Single) As Single
    TFY = (Y * zoom) + offy
End Function

Private Function TFXi(ByVal X As Single) As Single
    TFXi = (X - offx) / zoom
End Function

Private Function TFYi(ByVal Y As Single) As Single
    TFYi = (Y - offy) / zoom
End Function

Private Sub tlbTools_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.key
    Case "select"
        toolmode = tool_select
        picMain.MousePointer = vbDefault
    Case "move"
        toolmode = tool_move
        picMain.MousePointer = vbSizeAll
    Case "scale"
        toolmode = tool_scale
        picMain.MousePointer = vbSizeNESW
    Case "uvxneg"
        MoveVerts -1, 0
        picMain_Paint
        frmMain.picMain_Paint
        Me.picMain.SetFocus
    Case "uvxpos"
        MoveVerts 1, 0
        picMain_Paint
        frmMain.picMain_Paint
        Me.picMain.SetFocus
    Case "uvyneg"
        MoveVerts 0, -1
        picMain_Paint
        frmMain.picMain_Paint
        Me.picMain.SetFocus
    Case "uvypos"
        MoveVerts 0, 1
        picMain_Paint
        frmMain.picMain_Paint
        Me.picMain.SetFocus
    End Select
End Sub
