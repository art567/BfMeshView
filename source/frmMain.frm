VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "BfMeshView"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   13260
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   573
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   884
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrTime 
      Enabled         =   0   'False
      Left            =   8520
      Top             =   4560
   End
   Begin MSComctlLib.ImageList imlAnim 
      Left            =   1080
      Top             =   4920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1042
            Key             =   "rewind"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1394
            Key             =   "end"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":16E6
            Key             =   "pause"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1A38
            Key             =   "stop"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1D8A
            Key             =   "play"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picAnim 
      BorderStyle     =   0  'None
      HasDC           =   0   'False
      Height          =   495
      Left            =   1440
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   697
      TabIndex        =   6
      Top             =   6960
      Width           =   10455
      Begin VB.PictureBox picTime 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         FillStyle       =   0  'Solid
         Height          =   300
         Left            =   1800
         ScaleHeight     =   20
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   441
         TabIndex        =   8
         Top             =   0
         Width           =   6615
      End
      Begin MSComctlLib.Toolbar tlbAnim 
         Height          =   330
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         ToolTips        =   0   'False
         Style           =   1
         ImageList       =   "imlAnim"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   4
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "start"
               Object.ToolTipText     =   "To Start"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "play"
               Object.ToolTipText     =   "Play/Pause"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "stop"
               Object.ToolTipText     =   "Stop"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "end"
               Object.ToolTipText     =   "To End"
               ImageIndex      =   2
            EndProperty
         EndProperty
      End
      Begin VB.Label labTime 
         Alignment       =   2  'Center
         Caption         =   "0/0"
         Height          =   255
         Left            =   8640
         TabIndex        =   9
         Top             =   0
         Width           =   690
      End
   End
   Begin VB.TextBox txtConsole 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3600
      TabIndex        =   5
      Top             =   6000
      Visible         =   0   'False
      Width           =   3975
   End
   Begin MSComctlLib.ImageList imlTree 
      Left            =   960
      Top             =   2640
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   13
      ImageHeight     =   13
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":20DC
            Key             =   "lod"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2266
            Key             =   "mat"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":23F0
            Key             =   "tex"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2532
            Key             =   "geom"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":267C
            Key             =   "trinum"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":27CE
            Key             =   "shader"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2950
            Key             =   "prop"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2A8A
            Key             =   "texmissing"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2BD0
            Key             =   "badlod"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2D46
            Key             =   "file"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView trvMain 
      Height          =   1335
      Left            =   480
      TabIndex        =   4
      Top             =   720
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   2355
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   0
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "imlTree"
      Appearance      =   1
   End
   Begin VB.PictureBox picSamples 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   7200
      ScaleHeight     =   73
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   97
      TabIndex        =   3
      Top             =   600
      Visible         =   0   'False
      Width           =   1455
   End
   Begin ComctlLib.StatusBar stsMain 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   2
      Top             =   8310
      Width           =   13260
      _ExtentX        =   23389
      _ExtentY        =   503
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   6
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   13229
            Key             =   "info"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   1323
            MinWidth        =   1323
            Key             =   "geom"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   1323
            MinWidth        =   1323
            Key             =   "lod"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   2117
            MinWidth        =   2117
            Key             =   "mat"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   2117
            MinWidth        =   2117
            Key             =   "tri"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel6 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   2646
            MinWidth        =   2646
            Key             =   "mem"
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtLog 
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   3120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   2760
      Visible         =   0   'False
      Width           =   3975
   End
   Begin MSComDlg.CommonDialog cdlFile 
      Left            =   7800
      Top             =   3000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picMain 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   1575
      Left            =   3720
      OLEDropMode     =   2  'Automatic
      ScaleHeight     =   105
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   161
      TabIndex        =   1
      Top             =   480
      Width           =   2415
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuFileOpen 
         Caption         =   "Open..."
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "Save"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save As..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuFileReload 
         Caption         =   "Reload"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "Close"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuFileLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExport 
         Caption         =   "Export..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuFileLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "View"
      Begin VB.Menu mnuViewWireframe 
         Caption         =   "Wireframe"
      End
      Begin VB.Menu mnuViewLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewLighting 
         Caption         =   "Lighting"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewTexture 
         Caption         =   "Textures"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewLine3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewPolygons 
         Caption         =   "Polygons"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewVertices 
         Caption         =   "Vertices"
      End
      Begin VB.Menu mnuViewEdges 
         Caption         =   "Edges"
      End
      Begin VB.Menu mnuViewNormals 
         Caption         =   "Normals"
      End
      Begin VB.Menu mnuViewTangents 
         Caption         =   "Tangents"
      End
      Begin VB.Menu mnuViewBackfaces 
         Caption         =   "Backfaces"
      End
      Begin VB.Menu mnuViewBounds 
         Caption         =   "Bounding Box"
      End
      Begin VB.Menu mnuViewBonesys 
         Caption         =   "Bone System"
      End
      Begin VB.Menu mnuViewSamples 
         Caption         =   "Samples"
      End
      Begin VB.Menu mnuViewLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewAxis 
         Caption         =   "Show Axis"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewGrids 
         Caption         =   "Show Grids"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewLine88 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewMode 
         Caption         =   "Advanced Rendering"
         Begin VB.Menu mnuViewModeNormal 
            Caption         =   "Normal"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuViewModeVertexOrder 
            Caption         =   "Vertex Order"
         End
         Begin VB.Menu mnuViewModeOverdraw 
            Caption         =   "Triangle Overdraw"
         End
      End
      Begin VB.Menu mnuViewLine5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewSamplesBitmap 
         Caption         =   "Samples Bitmap"
      End
      Begin VB.Menu mnuViewLog 
         Caption         =   "Debug Console"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "Tools"
      Begin VB.Menu mnuToolsReloadTextures 
         Caption         =   "Reload Textures"
      End
      Begin VB.Menu mnuToolsLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToolsRenderLighting 
         Caption         =   "Render Lighting..."
      End
      Begin VB.Menu mnuToolsConvertNormalMap 
         Caption         =   "Convert Normal Map..."
      End
      Begin VB.Menu mnuToolsGenSamples 
         Caption         =   "Generate Samples..."
      End
      Begin VB.Menu mnuToolsMoveVerts 
         Caption         =   "Vertex Transform..."
      End
      Begin VB.Menu mnuToolsRenameTextures 
         Caption         =   "Rename Textures..."
      End
      Begin VB.Menu mnuToolsLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToolsMakeWorldSpace 
         Caption         =   "Make Tangents Worldspace"
      End
      Begin VB.Menu mnuToolsVeggieNormals 
         Caption         =   "Compute Veggie Normals"
      End
      Begin VB.Menu mnuToolsFlattenSamples 
         Caption         =   "Flatten Samples"
      End
      Begin VB.Menu mnuToolsFixSamples 
         Caption         =   "Fix Samples"
      End
      Begin VB.Menu mnuToolsVerifyMesh 
         Caption         =   "Fix Mesh"
      End
      Begin VB.Menu mnuToolsFixTexPaths 
         Caption         =   "Fix Texture Paths"
      End
      Begin VB.Menu mnuToolsLine3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToolsUvEditor 
         Caption         =   "UV Editor..."
      End
      Begin VB.Menu mnuToolsSkin 
         Caption         =   "Skin..."
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Options"
      Begin VB.Menu mnuOptionsLoadTextures 
         Caption         =   "Load Textures"
      End
      Begin VB.Menu mnuOptionsLoadSamples 
         Caption         =   "Load Samples"
      End
      Begin VB.Menu mnuOptionsLoadCon 
         Caption         =   "Load Con"
      End
      Begin VB.Menu mnuOptionsLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOptionsRememberViewSettings 
         Caption         =   "Remember Settings"
      End
      Begin VB.Menu mnuOptionsResetSettings 
         Caption         =   "Reset Settings To Default"
      End
      Begin VB.Menu mnuOptionsLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOptionsPreferences 
         Caption         =   "Preferences..."
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuHelpOpenglInfo 
         Caption         =   "OpenGL Info..."
      End
      Begin VB.Menu mnuHelpLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "About..."
      End
   End
   Begin VB.Menu mnuTreeContext 
      Caption         =   "dummy"
      Visible         =   0   'False
      Begin VB.Menu mnuTreeContextViewTex 
         Caption         =   "View Texture"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuTreeContextOpenFolder 
         Caption         =   "Open Folder"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuTreeContextLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTreeContextEdit 
         Caption         =   "Edit..."
      End
      Begin VB.Menu mnuTreeContextLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTreeContextCopy 
         Caption         =   "Copy"
      End
      Begin VB.Menu mnuTreeContextPaste 
         Caption         =   "Paste"
      End
      Begin VB.Menu mnuTreeContextLine3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTreeContextExpand 
         Caption         =   "Expand All"
      End
      Begin VB.Menu mnuTreeContextCollapse 
         Caption         =   "Collapse All"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private hglrc As Long
Private viewport_w As Long
Private viewport_h As Long

Private restorefrmrender As Boolean
Private restorefrmrenderwinstate As Long

Private restorefrmoutput As Boolean
Private restorefrmoutputwinstate As Long

Private restorefrmuvedit As Boolean
Private restorefrmuveditwinstate As Long

Private fps As Long
Private idemode As Boolean
Private current_file As String
Private current_filterindex As Integer
Private treeview_width As Long
Private menu_height As Long
Private splitdrag As Boolean
Private mousex As Single
Private mousey As Single
Private animtime As Single

Private copymatset As Boolean
Private copymat As bf2_mat

Public quitguard As Boolean

Private treemouseup As Boolean
Private treemousebutton As Integer
Private treemousex As Single
Private treemousey As Single


'form load
Private Sub Form_Load()
Dim pfd As PIXELFORMATDESCRIPTOR
Dim fmt As Long
    
    treeview_width = 270
    'menu_height = 30
    menu_height = 0
    
    'set config filename
    app_configfile = App.path & "\config.ini"
    
    'setup opengl
    pfd.nSize = Len(pfd)
    pfd.nVersion = 1
    pfd.dwFlags = PFD_SUPPORT_OPENGL Or PFD_DRAW_TO_WINDOW Or PFD_DOUBLEBUFFER Or PFD_TYPE_RGBA
    pfd.iPixelType = PFD_TYPE_RGBA
    pfd.cColorBits = 24
    pfd.cDepthBits = 16
    pfd.cStencilBits = 8
    pfd.iLayerType = PFD_MAIN_PLANE
    fmt = ChoosePixelFormat(Me.picMain.hdc, pfd)
    If fmt = 0 Then
        MsgBox "OpenGL initalization failed.", vbCritical
        Exit Sub
    End If
    fmt = SetPixelFormat(Me.picMain.hdc, fmt, pfd)
    hglrc = wglCreateContext(Me.picMain.hdc)
    wglMakeCurrent Me.picMain.hdc, hglrc
    
    'default states
    glTexEnvi GL_TEXTURE_ENV, GL_TEXTURE_ENV_MODE, GL_MODULATE
    glHint GL_PERSPECTIVE_CORRECTION_HINT, GL_NICEST
    glLightModeli GL_LIGHT_MODEL_TWO_SIDE, GL_TRUE
    'glEnable GL_NORMALIZE 'don't enable, allows us to see normal defects in models!
    
    'detect error reporting driver optimization
    'glBegin GL_TRIANGLES
    '    glBindTexture GL_TEXTURE_2D, tex
    'glEnd
    'Dim gr As GLenum
    'gr = glGetError
    'Me.stsMain.SimpleText = gr
    
    'init glew
    'If glewInit = GLEW_OK Then
    '    'MsgBox "glew ok"
    '    'glMultiTexCoord2f GL_TEXTURE0, 0, 0
    'End If
    
    'init extensions
    glextInit
    
    'light and material
    Dim dif(3) As GLfloat
    Dim amb(3) As GLfloat
    'glLightModeli GL_LIGHT_MODEL_LOCAL_VIEWER, 1
    glLightModelfv GL_LIGHT_MODEL_Ambient, amb(0)
    dif(0) = 0.5: dif(1) = 0.5: dif(2) = 0.5
    amb(0) = 0.5: amb(1) = 0.5: amb(2) = 0.5
    glLightfv GL_LIGHT0, GL_DIFFUSE, dif(0)
    glLightfv GL_LIGHT0, GL_AMBIENT, amb(0)
    glEnable GL_LIGHT0
    glEnable GL_COLOR_MATERIAL
    
    'default camera rotation
    camrotx = 10
    camroty = -30
    camzoom = 10
    
    'generate random color table
    GenColorTable
    
    'clean up
    CloseFile
    current_folder = App.path
    current_filterindex = 17
    
    'default config
    LoadDefaultConfig
    
    'load config file
    LoadConfig app_configfile
    
    'synchronize menu
    SyncMenu
    
    'hook
    If IsIdeMode Then
        idemode = True
    Else
        Hook Me.hWnd
        DisableTreeViewToolTips Me.trvMain.hWnd
    End If
    
    'maximize
    If opt_runmaximized Then
        Me.WindowState = vbMaximized
    End If
    
    'show form
    Me.Show
    Me.Refresh
    
    'timer
    Me.tmrTime.Interval = 1000 / 60
    SetTime 0
    
    quitguard = True
    
    'command line
Dim cmd As String
    cmd = Replace(Command$(), Chr(34), "")
    If Len(cmd) Then
        OpenFile cmd
    End If
    
    quitguard = False
    
    'fix for Win7
    picMain_Paint
    
End Sub


'form unload
Private Sub Form_Unload(Cancel As Integer)
    
    If quitguard Then
        Cancel = True
        Exit Sub
    End If
    
    CloseFile
    
    SaveConfig app_configfile
    
    app_exit = True
    
    Unload frmRender
    Unload frmOutput
    Unload frmUvEdit
    Unload frmTransform
    Unload frmSkin
    
    If hglrc Then
        wglMakeCurrent 0, 0
        wglDeleteContext hglrc
        hglrc = 0
    End If
    If Not idemode Then
        UnHook
    End If
End Sub


'updates menu checkboxes
Public Sub SyncMenu()
    frmMain.mnuViewWireframe.Checked = view_wire
    frmMain.mnuViewLighting.Checked = view_lighting
    frmMain.mnuViewTexture.Checked = view_textures
    frmMain.mnuViewPolygons.Checked = view_poly
    frmMain.mnuViewVertices.Checked = view_verts
    frmMain.mnuViewEdges.Checked = view_edges
    frmMain.mnuViewNormals.Checked = view_normals
    frmMain.mnuViewTangents.Checked = view_tangents
    frmMain.mnuViewBackfaces.Checked = view_backfaces
    frmMain.mnuViewBounds.Checked = view_bounds
    frmMain.mnuViewBonesys.Checked = view_bonesys
    frmMain.mnuViewSamples.Checked = view_samples
    frmMain.mnuViewAxis.Checked = view_axis
    frmMain.mnuViewGrids.Checked = view_grids
    
    frmMain.mnuOptionsLoadTextures.Checked = opt_loadtextures
    frmMain.mnuOptionsLoadSamples.Checked = opt_loadsamples
    frmMain.mnuOptionsLoadCon.Checked = opt_loadcon
    frmMain.mnuOptionsRememberViewSettings.Checked = opt_loadviewsettings
End Sub


'form resize
Public Sub Form_Resize()
'On Error GoTo errorhandler
    If Me.WindowState = vbMinimized Then
        
        'hide windows
        If frmRender.Visible Then
            frmRender.Visible = False
            restorefrmrender = True
            restorefrmrenderwinstate = frmRender.WindowState
        End If
        If frmOutput.Visible Then
            frmOutput.Visible = False
            restorefrmoutput = True
            restorefrmoutputwinstate = frmOutput.WindowState
        End If
        If frmUvEdit.Visible Then
            frmUvEdit.Visible = False
            restorefrmuvedit = True
            restorefrmuveditwinstate = frmUvEdit.WindowState
        End If
        
    Else
        
        'restore windows
        If restorefrmrender Then
            frmRender.Visible = True
            DoEvents
            frmRender.WindowState = restorefrmrenderwinstate
            restorefrmrender = False
        End If
        If restorefrmoutput Then
            frmOutput.Visible = True
            DoEvents
            frmOutput.WindowState = restorefrmoutputwinstate
            restorefrmoutput = False
        End If
        If restorefrmuvedit Then
            frmUvEdit.Visible = True
            DoEvents
            frmUvEdit.WindowState = restorefrmuveditwinstate
            restorefrmuvedit = False
        End If
        
        If idemode Then
            If Me.width < 400 * 15 Then Me.width = 400 * 15
            If Me.height < 300 * 15 Then Me.height = 300 * 15
        End If
        
        Dim yoff As Long
        yoff = menu_height + 1
        
        Dim animHeight As Long
        If picAnim.Visible Then
            Me.picAnim.height = Me.tlbAnim.height
            animHeight = picAnim.height
        End If
        
        If Me.ScaleWidth - 20 < treeview_width Then treeview_width = Me.ScaleWidth - 20
        
        'If Not idemode Then LockWindowUpdate Me.hWnd
        Me.trvMain.Move 2, yoff, treeview_width, Me.ScaleHeight - 3 - Me.stsMain.height - yoff - animHeight
        Me.picMain.Move treeview_width + 7, yoff, Me.ScaleWidth - (treeview_width + 4) - 4, Me.ScaleHeight - 3 - Me.stsMain.height - yoff - animHeight
        Me.picSamples.Move picMain.Left, picMain.top, picMain.width, picMain.height
        
        Me.txtLog.Move treeview_width + 7, yoff, Me.ScaleWidth - (treeview_width + 4) - 4, Me.ScaleHeight - 3 - Me.stsMain.height - yoff - Me.txtConsole.height - 6 - animHeight
        Me.txtConsole.Move treeview_width + 7, yoff + Me.txtLog.height + 3, Me.ScaleWidth - (treeview_width + 4) - 4 - animHeight
        
        Me.picAnim.Move 2, Me.picMain.top + Me.picMain.height + 2, Me.ScaleWidth - 4
        
        'If Not idemode Then LockWindowUpdate 0
        
        Me.Refresh
        
        viewport_w = Me.picMain.ScaleWidth
        viewport_h = Me.picMain.ScaleHeight
        If viewport_w < 1 Then viewport_w = 1
        If viewport_h < 1 Then viewport_h = 1
        picMain_Paint
        
    End If
    Exit Sub
errorhandler:
    'nothing
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        If SplitHover(x, y) Then
            splitdrag = True
        End If
    End If
    mousex = x
    mousey = y
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If splitdrag Then
        treeview_width = treeview_width + (x - mousex)
        treeview_width = Clamp(treeview_width, 20, 3 * (Me.ScaleWidth / 4))
        Form_Resize
    End If
    
    'hover
    If SplitHover(x, y) Then
        Me.MousePointer = vbSizeWE
    Else
        Me.MousePointer = vbDefault
    End If
    
    mousex = x
    mousey = y
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    splitdrag = False
End Sub

Private Function SplitHover(ByVal x As Single, ByVal y As Single) As Boolean
Dim minx As Single
Dim maxx As Single
    minx = Me.trvMain.Left + Me.trvMain.width
    maxx = Me.picMain.Left

Dim miny As Single
Dim maxy As Single
    miny = menu_height
    maxy = Me.ScaleHeight - Me.stsMain.height
    
    If x > minx And x < maxx Then
        If y > menu_height And y < maxy Then
            SplitHover = True
        End If
    End If
End Function


'picMain paint event
Public Sub picMain_Paint()
    If lmrender Then Exit Sub
    If hglrc = 0 Then Exit Sub
    
    If viewport_w = 0 Then Exit Sub
    If viewport_h = 0 Then Exit Sub
    
    glViewport 0, 0, viewport_w, viewport_h
    camasp = viewport_w / viewport_h
    
    'draw scene
    DrawScene
    
    'flip buffers
    SwapBuffers Me.picMain.hdc
End Sub

'form key input
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyLeft
        fooz = fooz - 1
        Me.Caption = fooz
        picMain_Paint
    Case vbKeyRight
        fooz = fooz + 1
        Me.Caption = fooz
        picMain_Paint
    Case vbKeyZ
        ZoomExtends
        picMain_Paint
    Case vbKeyF11
        Dim i As Long
        Dim s As Single
        Dim off As Single
        s = InputBox("Scale:", "Rescale", 1)
        If s > 0 And s <> 1 Then
            With vmesh
                For i = 0 To .vertnum - 1
                    off = i * (.vertstride / 4)
                    .vert(off + 0) = .vert(off + 0) * s
                    .vert(off + 1) = .vert(off + 1) * s
                    .vert(off + 2) = .vert(off + 2) * s
                Next i
            End With
        End If
        picMain_Paint
    End Select
End Sub

'forward picMain key input to form
Private Sub picMain_KeyDown(KeyCode As Integer, Shift As Integer)
    Form_KeyDown KeyCode, Shift
End Sub

Private Sub Pan(ByVal x As Single, ByVal y As Single)
    campanx = campanx + ((x * 0.1) * camzoom)
    campany = campany - ((y * 0.1) * camzoom)
End Sub

Private Sub zoom(ByVal v As Single)
    camzoom = camzoom - (v * camzoom * 0.1)
    If camzoom < 0.001 Then camzoom = 0.001
End Sub

Private Sub picMain_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    mouse_down = True
    mouse_px = x
    mouse_py = y
End Sub

Private Sub picMain_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim v As float3
    If mouse_down = True Then
        If Button = 1 Then
            camrotx = camrotx + ((y - mouse_py) * 0.5)
            camroty = camroty + ((x - mouse_px) * 0.5)
        End If
        If Button = 2 Then
            zoom ((mouse_py - y) * 0.1)
        End If
        If Button = 4 Then
            Pan ((x - mouse_px) * 0.01), ((y - mouse_py) * 0.01)
        End If
        picMain_Paint
    End If
    Me.MousePointer = vbDefault
    mouse_px = x
    mouse_py = y
End Sub

Public Sub MouseWheel(ByVal MouseKeys As Long, ByVal value As Long, ByVal x As Long, ByVal y As Long)
    zoom value
    picMain_Paint
End Sub

Private Sub picMain_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    mouse_down = False
    mouse_px = x
    mouse_py = y
End Sub

'--- anim ---------------------------------------------------------------------------------------------------------------

Private Sub picAnim_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub
    
    picTime.Move Me.tlbAnim.width + 2, 2, picAnim.ScaleWidth - tlbAnim.width - labTime.width - 6, 18
    
    labTime.top = 4
    labTime.Left = picAnim.ScaleWidth - labTime.width
    
    picTime_Paint
End Sub

Private Sub picTime_Paint()
    Dim w As Long
    Dim h As Long
    Dim wt As Long
    
    w = picTime.width
    h = picTime.height
    
    'slider position
    wt = (w - 2) * animtime
    
    'white background
    picTime.FillColor = RGB(255, 255, 255)
    picTime.ForeColor = RGB(127, 127, 127)
    DrawRect picTime.hdc, 0, 0, w, h
    
    'time
    If wt > 0 Then
        'picTime.FillColor = RGB(200, 200, 200)
        'picTime.ForeColor = RGB(200, 200, 200)
        picTime.FillColor = RGB(140, 214, 213)
        picTime.ForeColor = RGB(140, 214, 213)
        DrawRect picTime.hdc, 1, 1, wt + 1, (h / 2)
        
        picTime.FillColor = RGB(41, 165, 173)
        picTime.ForeColor = RGB(41, 165, 173)
        DrawRect picTime.hdc, 1, (h / 2), wt + 1, h - 1
    End If
    
    'draw slider
    picTime.FillColor = RGB(200, 63, 63)
    picTime.ForeColor = RGB(200, 63, 63)
    DrawRect picTime.hdc, wt + 1 - 2, 0, wt + 1 + 2, h
    
End Sub

Public Sub SetTime(ByVal t As Single, Optional redraw As Boolean = True)
    animtime = t
    If animtime < 0 Then animtime = 0
    If animtime > 1 Then animtime = 1
    
    Dim frame As Long
    If bf2baf.loaded Then
        frame = (bf2baf.framenum - 1) * animtime
        Me.labTime = (frame + 1) & "/" & bf2baf.framenum
        Me.labTime.Refresh
        
        DeformBF2Skeleton bf2baf, frame
    End If
    
    picTime_Paint
    If redraw Then
        picMain_Paint
    End If
End Sub

Private Sub picTime_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    picTime_MouseMove Button, Shift, x, y
End Sub

Private Sub picTime_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button Then
        Dim val As Single
        val = (x - 1) / (picTime.width - 2)
        If val < 0 Then val = 0
        If val > 1 Then val = 1
        SetTime val
    End If
End Sub

Private Sub tmrTime_Timer()
    Dim frames As Single
    frames = 100
    If bf2baf.loaded Then frames = bf2baf.framenum
    
    Dim t As Single
    t = animtime + (1 / 60) * ((1 / CSng(frames)) * 30)
    
    'repeat
    If t > 1 Then t = 0
    
    SetTime t
End Sub

Private Sub tlbAnim_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.key
    Case "start"
        SetTime 0
    Case "end"
        SetTime 1
    Case "play"
        tmrTime.Enabled = Not tmrTime.Enabled
        If tmrTime.Enabled Then
            Button.Image = "pause"
        Else
            Button.Image = "play"
        End If
    Case "stop"
        tlbAnim.Buttons("play").Image = "play"
        SetTime 0, False
        tmrTime.Enabled = False
        ResetBF2Skeleton
        BF2MeshDeform
        frmMain.picMain_Paint
    End Select
End Sub

'--- menu ---------------------------------------------------------------------------------------------------------------

Private Sub mnuFileOpen_Click()
    With Me.cdlFile
        .DialogTitle = "Open File"
        .Filter = "Static Mesh (*.staticmesh)|*.staticmesh|" & _
                  "Bundled Mesh (*.bundledmesh)|*.bundledmesh|" & _
                  "Skin Mesh (*.skinnedmesh)|*.skinnedmesh|" & _
                  "Collision Mesh (*.collisionmesh)|*.collisionmesh|" & _
                  "BF2 Animation (*.baf)|*.baf|" & _
                  "Skeleton (*.ske)|*.ske|" & _
                  "Occluder (*.occ)|*.occ|" & _
                  "Samples (*.samp*)|*.samp*|" & _
                  "Standard Mesh (*.sm)|*.sm|" & _
                  "Tree Mesh (*.tm)|*.tm|" & _
                  "Wavefront OBJ (*.obj)|*.obj|" & _
                  "FrostBite Mesh (*.res)|*.res|" & _
                  "FHX Geometry (*.geo)|*.geo|" & _
                  "FHX Trimesh (*.tri)|*.tri|" & _
                  "FHX Rig (*.rig)|*.rig|" & _
                  "BfMeshView Workspace (*.bfmv)|*.bfmv|" & _
                  "All Files (*.*)|*.*"
        .FilterIndex = current_filterindex
        .flags = cdlOFNExplorer Or cdlOFNFileMustExist
        .InitDir = current_folder
        .CancelError = True
        On Error Resume Next
        .ShowOpen
        If Not err.Number = cdlCancel Then
            On Error GoTo 0
            
            current_filterindex = .FilterIndex
            OpenFile .filename
        End If
    End With
End Sub

Private Sub mnuFileSave_Click()
    If MsgBox("Are you sure you want to overwrite this file?", vbExclamation Or vbYesNoCancel) = vbYes Then
        SaveFile current_file
    End If
End Sub

Private Sub mnuFileSaveAs_Click()
    With Me.cdlFile
        .DialogTitle = "Save File"
        .Filter = "Static Mesh (*.staticmesh)|*.staticmesh|" & _
                  "Bundled Mesh (*.bundledmesh)|*.bundledmesh|" & _
                  "Skin Mesh (*.skinnedmesh)|*.skinnedmesh|" & _
                  "Samples (*.samp*)|*.samp*|" & _
                  "All Files (*.*)|*.*"
        .flags = cdlOFNOverwritePrompt Or cdlOFNNoReadOnlyReturn
        .InitDir = current_folder
        .CancelError = True
        
        If vmesh.loadok Then
            .filename = GetNameFromFileName(vmesh.filename)
            If vmesh.isBundledMesh Then
                .FilterIndex = 2
            ElseIf vmesh.isSkinnedMesh Then
                .FilterIndex = 3
            Else
                .FilterIndex = 1
            End If
        Else
            .filename = GetFilenameFromPath(bf2samples(0).filename)
            .FilterIndex = 4
        End If
        
        On Error Resume Next
        .ShowSave
        If Not err.Number = cdlCancel Then
            On Error GoTo 0
            
            SaveFile .filename
        End If
    End With
End Sub

Private Sub mnuFileReload_Click()
Dim oldfile As String
    oldfile = current_file
    
    CloseFile
    OpenFile oldfile
End Sub

Private Sub mnuFileClose_Click()
    CloseFile
End Sub

Private Sub mnuFileExport_Click()
    With Me.cdlFile
        .DialogTitle = "Export File"
        .Filter = "Wavefront OBJ (*.obj)|*.obj"
        .FilterIndex = 1
        .flags = cdlOFNOverwritePrompt Or cdlOFNNoReadOnlyReturn
        .InitDir = current_folder
        .CancelError = True
        .filename = GetNameFromFileName(vmesh.filename)
        On Error Resume Next
        .ShowSave
        If Not err.Number = cdlCancel Then
            On Error GoTo 0
            
            ExportMesh .filename
        End If
    End With
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuViewWireframe_Click()
    view_wire = Not view_wire
    mnuViewWireframe.Checked = view_wire
    picMain_Paint
End Sub

Private Sub mnuViewLighting_Click()
    view_lighting = Not view_lighting
    mnuViewLighting.Checked = view_lighting
    picMain_Paint
End Sub

Private Sub mnuViewTexture_Click()
    view_textures = Not view_textures
    mnuViewTexture.Checked = view_textures
    picMain_Paint
End Sub

Private Sub mnuViewPolygons_Click()
    view_poly = Not view_poly
    mnuViewPolygons.Checked = view_poly
    picMain_Paint
End Sub

Private Sub mnuViewVertices_Click()
    view_verts = Not view_verts
    mnuViewVertices.Checked = view_verts
    picMain_Paint
End Sub

Public Sub mnuViewEdges_Click()
    view_edges = Not view_edges
    mnuViewEdges.Checked = view_edges
    picMain_Paint
End Sub

Private Sub mnuViewNormals_Click()
    view_normals = Not view_normals
    mnuViewNormals.Checked = view_normals
    picMain_Paint
End Sub

Private Sub mnuViewBackfaces_Click()
    view_backfaces = Not view_backfaces
    mnuViewBackfaces.Checked = view_backfaces
    picMain_Paint
End Sub

Private Sub mnuViewTangents_Click()
    view_tangents = Not view_tangents
    mnuViewTangents.Checked = view_tangents
    picMain_Paint
End Sub

Private Sub mnuViewBounds_Click()
    view_bounds = Not view_bounds
    mnuViewBounds.Checked = view_bounds
    picMain_Paint
End Sub

Private Sub mnuViewBonesys_Click()
    view_bonesys = Not view_bonesys
    mnuViewBonesys.Checked = view_bonesys
    picMain_Paint
End Sub

Private Sub mnuViewSamples_Click()
    view_samples = Not view_samples
    mnuViewSamples.Checked = view_samples
    picMain_Paint
End Sub

Private Sub mnuViewAxis_Click()
    view_axis = Not view_axis
    mnuViewAxis.Checked = view_axis
    picMain_Paint
End Sub

Private Sub mnuViewGrids_Click()
    view_grids = Not view_grids
    mnuViewGrids.Checked = view_grids
    picMain_Paint
End Sub

Private Sub mnuViewModeNormal_Click()
    draw_mode = dm_normal
    Me.mnuViewModeNormal.Checked = True
    Me.mnuViewModeVertexOrder.Checked = False
    Me.mnuViewModeOverdraw.Checked = False
    picMain_Paint
End Sub

Private Sub mnuViewModeVertexOrder_Click()
    draw_mode = dm_vertorder
    Me.mnuViewModeNormal.Checked = False
    Me.mnuViewModeVertexOrder.Checked = True
    Me.mnuViewModeOverdraw.Checked = False
    picMain_Paint
End Sub

Private Sub mnuViewModeOverdraw_Click()
    draw_mode = dm_overdraw
    Me.mnuViewModeNormal.Checked = False
    Me.mnuViewModeVertexOrder.Checked = False
    Me.mnuViewModeOverdraw.Checked = True
    picMain_Paint
End Sub

Private Sub mnuViewSamplesBitmap_Click()
    mnuViewSamplesBitmap.Checked = Not mnuViewSamplesBitmap.Checked
    'Me.txtLog.Visible = Not mnuViewSamplesBitmap.Checked
    Me.picMain.Visible = Not mnuViewSamplesBitmap.Checked
    Me.picSamples.Visible = mnuViewSamplesBitmap.Checked
    
    If mnuViewSamplesBitmap.Checked Then
        DrawSamples2d Me.picSamples
    End If
End Sub

Private Sub mnuViewLog_Click()
    mnuViewLog.Checked = Not mnuViewLog.Checked
    Me.txtLog.Visible = mnuViewLog.Checked
    Me.picMain.Visible = Not mnuViewLog.Checked
    'Me.picSamples.Visible = Not mnuViewLog.Checked
    
    Me.txtConsole.Visible = Me.txtLog.Visible
End Sub

Private Sub mnuOptionsLoadTextures_Click()
    opt_loadtextures = Not opt_loadtextures
    mnuOptionsLoadTextures.Checked = opt_loadtextures
End Sub

Private Sub mnuOptionsLoadSamples_Click()
    opt_loadsamples = Not opt_loadsamples
    mnuOptionsLoadSamples.Checked = opt_loadsamples
End Sub

Private Sub mnuOptionsLoadCon_Click()
    opt_loadcon = Not opt_loadcon
    mnuOptionsLoadCon.Checked = opt_loadcon
End Sub

Private Sub mnuOptionsRememberViewSettings_Click()
    opt_loadviewsettings = Not opt_loadviewsettings
    mnuOptionsRememberViewSettings.Checked = opt_loadviewsettings
End Sub

Private Sub mnuOptionsResetSettings_Click()
    LoadDefaultConfig
    SyncMenu
    picMain_Paint
End Sub

Private Sub mnuOptionsPreferences_Click()
    frmOptions.Show vbModal
    picMain_Paint
End Sub

Public Sub mnuToolsReloadTextures_Click()
    LoadMeshTextures
    FillTreeView
    picMain_Paint
    SetStatus "info", "Done."
    
    'todo: update icons in treeview
End Sub

Private Sub mnuToolsRenderLighting_Click()
    frmRender.Show 'vbModal
End Sub

Private Sub mnuToolsConvertNormalMap_Click()
    frmNormalMap.Show vbModal
End Sub

Private Sub mnuToolsMakeWorldSpace_Click()
    If Not vmesh.loadok Then
        MsgBox "No BF2 mesh loaded.", vbExclamation
        Exit Sub
    End If
    VisMeshTool_MakeWS
    picMain_Paint
End Sub

Private Sub mnuToolsVeggieNormals_Click()
    If Not vmesh.loadok Then
        MsgBox "No BF2 mesh loaded.", vbExclamation
        Exit Sub
    End If
    VisMeshTool_VeggieNormals
    picMain_Paint
End Sub

Private Sub mnuToolsFlattenSamples_Click()
    If Not bf2samples(0).loaded Then
        MsgBox "No samples file loaded.", vbExclamation
        Exit Sub
    End If
    FlattenSamples bf2samples(0)
    picMain_Paint
End Sub

Private Sub mnuToolsFixSamples_Click()
    If Not bf2samples(0).loaded Then
        MsgBox "No samples file loaded.", vbExclamation
        Exit Sub
    End If
    FixSamples bf2samples(0)
    picMain_Paint
End Sub

Private Sub mnuToolsVerifyMesh_Click()
    BF2VerifyMesh
End Sub

Private Sub mnuToolsFixTexPaths_Click()
    BF2MeshFixTexPaths
End Sub

Private Sub mnuToolsGenSamples_Click()
    If Not vmesh.loadok Then
        MsgBox "No BF2 mesh loaded.", vbExclamation
        Exit Sub
    End If
    frmSamples.Show vbModal
End Sub

Private Sub mnuToolsMoveVerts_Click()
    If Not vmesh.loadok Then
        MsgBox "No BF2 mesh loaded.", vbExclamation
        Exit Sub
    End If
    frmTransform.Show
    picMain_Paint
End Sub

Private Sub mnuToolsRenameTextures_Click()
    If Not vmesh.loadok Then
        MsgBox "No BF2 mesh loaded.", vbExclamation
        Exit Sub
    End If
    frmTexRename.Show vbModal
End Sub

Private Sub mnuToolsUvEditor_Click()
    If Not vmesh.loadok Then
        MsgBox "No BF2 mesh loaded.", vbExclamation
        Exit Sub
    End If
    frmUvEdit.Show
End Sub

Private Sub mnuToolsSkin_Click()
    If Not vmesh.loadok And vmesh.isSkinnedMesh Then
        MsgBox "No BF2 mesh loaded.", vbExclamation
        Exit Sub
    End If
    frmSkin.Show
End Sub

Private Sub mnuHelpOpenglInfo_Click()
    frmGLInfo.Show vbModal
End Sub

Private Sub mnuHelpAbout_Click()
    frmAbout.Show vbModal
End Sub

'--- treeview menu -----------------------------------------------------------

Private Sub mnuTreeContextViewTex_Click()
    Dim fname As String
    If vmesh.loadok Then fname = GetSelectedTextureFilename
    If stdmesh.loadok Then fname = GetStdMeshSelectedTextureFilename
    'If treemesh.loadok Then fname = GetSelectedTextureFilename2
    
    If FileExist(fname) Then
        ShellExecute 0, vbNullString, fname, vbNullString, vbNullString, vbNormalFocus
    End If
End Sub

Private Sub mnuTreeContextOpenFolder_Click()
    Dim fname As String
    If vmesh.loadok Then
        fname = GetFilePath(GetSelectedTextureFilename)
        If FileExist(fname) Then
            ShellExecute 0, vbNullString, fname, vbNullString, vbNullString, vbNormalFocus
        End If
    End If
End Sub

Private Sub mnuTreeContextEdit_Click()
'Dim str As String
'Dim def As String
'    def = GetSelectedTextureFilename
'    str = InputBox("Texture:", "Edit Texture File Name", def)
'    If Len(str) > 0 Then
'        'setselectedtexturefilename str
'    End If
    
    If selmat > -1 Then
        frmMaterial.Show vbModal
        Exit Sub
    End If
End Sub

Private Function ValSel(ByVal level As Long) As Boolean
    With vmesh
        If Not .loadok Then Exit Function
        
        If selgeom < 0 Then Exit Function
        If selgeom > vmesh.geomnum - 1 Then Exit Function
        If level = 1 Then
            ValSel = True
            Exit Function
        End If
        
        If sellod < 0 Then Exit Function
        If sellod > vmesh.geom(selgeom).lodnum - 1 Then Exit Function
        If level = 2 Then
            ValSel = True
            Exit Function
        End If
        
        If selmat < 0 Then Exit Function
        If selmat > vmesh.geom(selgeom).lod(sellod).matnum Then Exit Function
        If level = 3 Then
            ValSel = True
            Exit Function
        End If
        
        If seltex < 0 Then Exit Function
        If seltex > vmesh.geom(selgeom).lod(sellod).mat(selmat).mapnum Then Exit Function
        If level = 4 Then
            ValSel = True
            Exit Function
        End If
        
    End With
End Function

Private Sub mnuTreeContextCopy_Click()
    If Not ValSel(3) Then Exit Sub
    copymat = vmesh.geom(selgeom).lod(sellod).mat(selmat)
    copymatset = True
End Sub

Private Sub mnuTreeContextPaste_Click()
    If Not ValSel(3) Then Exit Sub
    If Not copymatset Then Exit Sub
    PasteMaterial vmesh.geom(selgeom).lod(sellod).mat(selmat), copymat
    frmMain.FillTreeView
    picMain_Paint
End Sub

Private Sub mnuTreeContextExpand_Click()
Dim i As Long
    trvMain.Visible = False
    For i = 1 To trvMain.Nodes.count
        trvMain.Nodes.Item(i).Expanded = True
    Next i
    trvMain.Visible = True
End Sub

Private Sub mnuTreeContextCollapse_Click()
Dim i As Long
    trvMain.Visible = False
    For i = 1 To trvMain.Nodes.count
        trvMain.Nodes.Item(i).Expanded = False
    Next i
    trvMain.Visible = True
End Sub

Private Function GetSelectedTextureFilename() As String
    GetSelectedTextureFilename = BF2GetTextureFilename(selgeom, sellod, selmat, seltex)
End Function

Private Sub trvMain_DblClick()
    If seltex > -1 Then
        mnuTreeContextViewTex_Click
    End If
End Sub

Private Sub trvMain_KeyUp(KeyCode As Integer, Shift As Integer)
    SelectMesh val(trvMain.SelectedItem.tag), trvMain.SelectedItem.key
    picMain_Paint
End Sub

Private Sub trvMain_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error GoTo hell
    
    'select node under cursor
    Dim n As MSComctlLib.node
    Set n = Me.trvMain.HitTest(x, y)
    If Not n Is Nothing Then
        SelectMesh n.tag, n.key
        picMain_Paint
    End If
    
    Exit Sub
hell:
    SelectMesh -1, MakeKey(0, 0, -1, -1)
    picMain_Paint
End Sub

Private Sub trvMain_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error GoTo errhandler
    
    'note: we show the context menu on MouseDown because NodeClick steals the MouseUp event
    If Button = vbRightButton Then
        
        Me.mnuTreeContextViewTex.Enabled = False
        Me.mnuTreeContextOpenFolder.Enabled = False
        Me.mnuTreeContextEdit.Enabled = False
        Me.mnuTreeContextCopy.Enabled = False
        Me.mnuTreeContextPaste.Enabled = False
        
        If vmesh.loadok Then
            If selmat > -1 And seltex > -1 Then
                If GetSelectedTextureFilename <> "" Then
                    Me.mnuTreeContextViewTex.Enabled = True
                    Me.mnuTreeContextOpenFolder.Enabled = True
                End If
            End If
            
            If selmat > -1 Then
                Me.mnuTreeContextEdit.Enabled = True
            End If
            If selmat > -1 And seltex = -1 Then
                Me.mnuTreeContextCopy.Enabled = True
                Me.mnuTreeContextPaste.Enabled = True
            End If
        End If
        
        'show context menu
        PopupMenu Me.mnuTreeContext, vbPopupMenuRightButton, _
                  trvMain.Left + (x / 15) + 2, _
                  trvMain.top + (y / 15) + 2, mnuTreeContextViewTex
        
    End If
    
    Exit Sub
errhandler:
    MsgBox "trvMain_MouseUp" & vbLf & err.description, vbCritical
End Sub

'NOTE: DONT USE THIS, BREAKS MouseUp and MouseDown events!!!
'Private Sub trvMain_NodeClick(ByVal node As MSComctlLib.node)
'    If node.tag <> "" Then
'        SelectMesh node.tag, node.key
'        picMain_Paint
'    End If
'End Sub

Public Sub SelectMesh(ByVal id As Long, Optional ByVal key As String)
    'On Error Resume Next
    
    If vmesh.loadok Then
        
        're-enable drawing to allow some form of recovery
        vmesh.drawok = True
        
        'reset selection in case of error
        selgeom = 0
        sellod = 0
        selmat = -1
        seltex = -1
        
        If id = -1 Then
            Dim str() As String
            
            str = Split(key, "|")
            If str(0) = "@" Then
                selgeom = val(str(1))
                sellod = val(str(2))
                selmat = val(str(3))
                seltex = val(str(4))
            End If
        Else
            'old selection mechanism
            selgeom = GetBit(id, 0)
            sellod = GetBit(id, 1)
        End If
        
        'make sure the indices are not out of range
        If selgeom < 0 Then selgeom = 0
        If sellod < 0 Then sellod = 0
        If selmat < -1 Then selmat = -1
        If seltex < -1 Then seltex = -1
        
        If selgeom > vmesh.geomnum - 1 Then selgeom = 0
        If sellod > vmesh.geom(selgeom).lodnum - 1 Then sellod = 0
        If selmat > 0 Then
            If selmat > vmesh.geom(selgeom).lod(sellod).matnum Then selmat = -1
        End If
        If seltex > 0 Then
            If seltex > vmesh.geom(selgeom).lod(sellod).mat(selmat).mapnum Then seltex = -1
        End If
        
        'update status bar
        SetStatus "geom", "Geom " & selgeom
        SetStatus "lod", "Lod " & sellod
        SetStatus "mat", "Material " & selmat
        SetStatus "tri", vmesh.geom(selgeom).lod(sellod).polycount & " triangles"
        
        'redraw UV editor
        UvEdit_Paint
    End If
    
    If cmesh.loadok Then
        selgeom = GetBit(id, 0)
        selsub = GetBit(id, 1)
        sellod = GetBit(id, 2)
        selmat = -1
        
        If selgeom > cmesh.geomnum - 1 Then selgeom = 0
        If selsub > cmesh.geom(selgeom).subgnum - 1 Then selsub = 0
        If sellod > cmesh.geom(selgeom).subg(selsub).lodnum - 1 Then sellod = 0
        
        SetStatus "geom", "Geom " & selgeom
        SetStatus "lod", "Sub " & selsub
        SetStatus "mat", "Col " & sellod
        SetStatus "tri", ""
    End If
    
    If stdmesh.loadok Then
        Dim b1 As Byte
        b1 = GetBit(id, 0)
        
        If b1 < 2 Then
            selgeom = GetBit(id, 0)
            sellod = GetBit(id, 1)
            selmat = GetBit(id, 2)
            seltex = -1
        End If
        
        SetStatus "geom", ""
        If b1 = 0 Then
            SetStatus "lod", "Lod " & sellod
            SetStatus "mat", stdmesh.lod(sellod).matnum & " materials"
            SetStatus "tri", stdmesh.lod(sellod).polycount & " triangles"
        End If
        If b1 = 1 Then
            SetStatus "lod", "Col " & sellod
            SetStatus "mat", ""
            SetStatus "tri", stdmesh.col(sellod).facenum & " triangles"
        End If
        If b1 = 2 Then 'we abuse seltex to know which shader is selected
            seltex = GetBit(id, 1)
        End If
    End If
    
    If fhxgeo.loadok Then
        selgeom = -1
        sellod = id
        selmat = -1
        
        If sellod < 0 Then sellod = 0
        If sellod > fhxgeo.lodnum - 1 Then sellod = fhxgeo.lodnum - 1
        
        SetStatus "geom", ""
        SetStatus "lod", "Lod " & sellod
        SetStatus "mat", (fhxgeo.lod(sellod).matgroupnum) & " materials"
        SetStatus "tri", (fhxgeo.lod(sellod).indexnum / 3) & " triangles"
    End If
    
    On Error GoTo 0
End Sub

Private Sub cmdLog_Click()
    mnuViewLog_Click
End Sub

'--- files --------------------------------------------------------------------------

'drop file on form
Public Sub DropFile(ByRef str As String)
    If FileExist(str) Then
        Select Case LCase(GetFileExt(str))
        Case "baf"
            If LoadBF2Anim(str) Then
                SetTime 0
            End If
        Case "ske"
            LoadBF2Skeleton str
        Case "con", "tweak"
            LoadCon str
        Case Else
            OpenFile str
        End Select
    End If
    FillTreeView
    Form_Resize
    picMain_Paint
End Sub

Private Sub OpenFile(ByVal filename As String)
    CloseFile
    
    copymatset = False
    seldefault = MakeTag(0, 0, 0)
    
    current_file = filename
    current_folder = GetFilePath(filename)
    
    OpenMeshFile filename
    FillTreeView
    'ZoomExtends
    
    SetStatus "info", "Done."
    SetStatus "mem", "TexMem: " & GetTextureMemory
    
    'reset selection
    SelectMesh seldefault
    
    mnuFileSave.Enabled = True
    mnuFileSaveAs.Enabled = True
    mnuFileClose.Enabled = True
    mnuFileReload.Enabled = True
    mnuFileExport.Enabled = True
    
    UpdateCaption
    'picMain_Paint
    Form_Resize
    UvEdit_Paint
End Sub

Private Sub SaveFile(ByVal filename As String)
    current_file = filename
    
    current_folder = GetFilePath(filename)
    SaveMeshFile filename
    
    SetStatus "info", "Saved."
    
    UpdateCaption
End Sub

Private Sub CloseFile()
    current_file = ""
    
    CloseMeshFile
    FillTreeView
    txtLog.Text = ""
    
    SetStatus "info", ""
    SetStatus "geom", ""
    SetStatus "lod", ""
    SetStatus "mat", ""
    SetStatus "mem", ""
    
    mnuFileSave.Enabled = False
    mnuFileSaveAs.Enabled = False
    mnuFileClose.Enabled = False
    mnuFileReload.Enabled = False
    mnuFileExport.Enabled = False
    
    UpdateCaption
    'picMain_Paint
    Form_Resize
    UvEdit_Paint
End Sub

'updates title bar
Private Sub UpdateCaption()
    If Len(current_file) = 0 Then
        frmMain.Caption = App.Title & " " & App.Major & "." & App.Minor & "." & App.Revision
    Else
        frmMain.Caption = App.Title & " " & App.Major & "." & App.Minor & "." & App.Revision & " - [" & current_file & "]"
    End If
End Sub

'reset mousepointer
Private Sub stsMain_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Me.MousePointer = vbDefault
End Sub
Private Sub trvMain_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Me.MousePointer = vbDefault
End Sub

Private Sub txtLog_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Me.MousePointer = vbDefault
End Sub

Private Sub txtConsole_KeyDown(KeyCode As Integer, Shift As Integer)
    If Not KeyCode = vbKeyReturn Then Exit Sub
    
    Dim cmd As String
    cmd = Me.txtConsole.Text
    Me.txtConsole.Text = ""
    
    Dim str() As String
    str() = Split(cmd, " ", 1)
    
    Select Case str(0)
    Case "weights"
        With vmesh
            Dim vw As bf2vw
            Dim i As Long
            Dim woff As Long
            Dim stride As Long
            woff = BF2MeshGetWeightOffset
            stride = .vertstride / 4
            
            Dim minb As Long
            Dim maxb As Long
            Dim minw As Long
            Dim maxw As Long
            minb = 999
            maxb = -999
            
            For i = 0 To .vertnum - 1
                CopyMem VarPtr(vw), VarPtr(.vert(i * stride + woff)), 4
                
                If i < 100 Then
                    Echo "vert[" & i & "].b1: " & vw.b1
                    Echo "vert[" & i & "].b2: " & vw.b2
                    Echo "vert[" & i & "].w1: " & vw.w1
                    Echo "vert[" & i & "].w2: " & vw.w2
                End If
                
                minb = min(minb, vw.b1)
                minb = min(minb, vw.b2)
                maxb = max(maxb, vw.b1)
                maxb = max(maxb, vw.b2)
                
                minw = min(minw, vw.w1)
                minw = min(minw, vw.w2)
                maxw = max(maxw, vw.w1)
                maxw = max(maxw, vw.w2)
            Next i
            
            Echo "bone range: " & minb & " to " & maxb
            Echo "weight range: " & minw & " to " & maxw
        End With
    Case Else
        Echo "Unknown command."
    End Select
    
    Me.txtLog.SelStart = Len(Me.txtLog.Text)
    Exit Sub
errhandler:
    Echo "ERROR: " & err.description
    Me.txtLog.SelStart = Len(Me.txtLog.Text)
End Sub

'fills treeview
Public Sub FillTreeView()
    
    Me.trvMain.Visible = False
    Me.trvMain.Enabled = False
    
    'make backup of node states
    Dim i As Long
    Dim j As Long
    Dim n_num As Long
    Dim n_key() As String
    Dim n_exp() As Boolean
    Dim selkey As String
    If Me.trvMain.Nodes.count > 0 Then
        n_num = Me.trvMain.Nodes.count
        ReDim n_key(1 To n_num)
        ReDim n_exp(1 To n_num)
        For i = 1 To trvMain.Nodes.count
            n_key(i) = trvMain.Nodes.Item(i).key
            n_exp(i) = trvMain.Nodes.Item(i).Expanded
        Next i
        If Not Me.trvMain.SelectedItem Is Nothing Then
            selkey = Me.trvMain.SelectedItem.key
        End If
    End If
    
    'clear treeview
    Me.trvMain.Nodes.Clear
    
    'BF2
    FillTreeVisMesh Me.trvMain
    FillTreeColMesh Me.trvMain
    FillTreeBF2Skeleton Me.trvMain
    FillTreeBF2Anim Me.trvMain
    FillTreeSamples Me.trvMain
    FillTreeOcc Me.trvMain
    FillTreeBF2Con Me.trvMain
    
    'bf1942
    FillTreeStdMesh Me.trvMain
    FillTreeStdMeshShaders Me.trvMain
    FillTreeTreeMesh Me.trvMain
    
    'FHX
    FillTreeFhxGeo Me.trvMain
    FillTreeFhxTri Me.trvMain
    FillTreeFhxRig Me.trvMain
    
    'obj
    FillTreeObj Me.trvMain
    
    'restore node states
    If Me.trvMain.Nodes.count > 0 Then
        Dim delid As Long
        For i = 1 To trvMain.Nodes.count
            delid = 0
            For j = 1 To n_num
                'If trvMain.Nodes(i).key = n_key(j) Then
                If StrMatch(trvMain.Nodes(i).key, n_key(j)) Then
                    trvMain.Nodes.Item(i).key = n_key(j)
                    trvMain.Nodes.Item(i).Expanded = n_exp(j)
                    delid = j
                    Exit For
                End If
            Next j
            
            'shrink list
            If delid > 0 Then
                'n_key(delid) = n_key(n_num)
                'n_num = n_num - 1
            End If
            
            If LenB(selkey) > 0 Then
                If trvMain.Nodes(i).key = selkey Then
                    Set trvMain.SelectedItem = trvMain.Nodes(i)
                End If
            End If
        Next i
    End If
    
    Me.trvMain.Enabled = True
    Me.trvMain.Visible = True
End Sub

'fast string compare
Private Function StrMatch(ByRef a As String, ByRef b As String) As Boolean
    If LenB(a) = 0 Then Exit Function
    If LenB(b) = 0 Then Exit Function
    If AscW(a) <> AscW(b) Then Exit Function
    If a <> b Then Exit Function
    StrMatch = True
End Function

'redraws UV editor, but only if it's form is loaded
Private Sub UvEdit_Paint()
    If uveditor_isloaded Then
        frmUvEdit.picMain_Paint
    End If
End Sub
