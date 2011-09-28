VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMaterial 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit Material"
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8415
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMaterial.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   297
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   561
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Height          =   285
      Index           =   6
      Left            =   7920
      TabIndex        =   28
      Top             =   3360
      Width           =   330
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Height          =   285
      Index           =   5
      Left            =   7920
      TabIndex        =   27
      Top             =   3000
      Width           =   330
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Height          =   285
      Index           =   4
      Left            =   7920
      TabIndex        =   26
      Top             =   2640
      Width           =   330
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Height          =   285
      Index           =   3
      Left            =   7920
      TabIndex        =   25
      Top             =   2280
      Width           =   330
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Height          =   285
      Index           =   2
      Left            =   7920
      TabIndex        =   24
      Top             =   1920
      Width           =   330
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Height          =   285
      Index           =   1
      Left            =   7920
      TabIndex        =   23
      Top             =   1560
      Width           =   330
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Height          =   285
      Index           =   0
      Left            =   7920
      TabIndex        =   22
      Top             =   1200
      Width           =   330
   End
   Begin VB.TextBox txtMap 
      Height          =   285
      Index           =   6
      Left            =   1320
      TabIndex        =   20
      Top             =   3360
      Width           =   6495
   End
   Begin VB.TextBox txtMap 
      Height          =   285
      Index           =   5
      Left            =   1320
      TabIndex        =   18
      Top             =   3000
      Width           =   6495
   End
   Begin VB.TextBox txtTechnique 
      Height          =   285
      Left            =   1320
      TabIndex        =   11
      Top             =   840
      Width           =   3135
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   3960
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   7200
      TabIndex        =   9
      Top             =   3960
      Width           =   1095
   End
   Begin VB.TextBox txtShader 
      Height          =   285
      Left            =   1320
      TabIndex        =   7
      Top             =   480
      Width           =   3135
   End
   Begin VB.ComboBox cbbTransparency 
      Height          =   315
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   90
      Width           =   3135
   End
   Begin VB.TextBox txtMap 
      Height          =   285
      Index           =   4
      Left            =   1320
      TabIndex        =   5
      Top             =   2640
      Width           =   6495
   End
   Begin VB.TextBox txtMap 
      Height          =   285
      Index           =   3
      Left            =   1320
      TabIndex        =   4
      Top             =   2280
      Width           =   6495
   End
   Begin VB.TextBox txtMap 
      Height          =   285
      Index           =   2
      Left            =   1320
      TabIndex        =   3
      Top             =   1920
      Width           =   6495
   End
   Begin VB.TextBox txtMap 
      Height          =   285
      Index           =   1
      Left            =   1320
      TabIndex        =   2
      Top             =   1560
      Width           =   6495
   End
   Begin VB.TextBox txtMap 
      Height          =   285
      Index           =   0
      Left            =   1320
      TabIndex        =   1
      Top             =   1200
      Width           =   6495
   End
   Begin MSComDlg.CommonDialog cdlFile 
      Left            =   7800
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label labMisc 
      AutoSize        =   -1  'True
      Caption         =   "Texture 7:"
      Height          =   195
      Index           =   9
      Left            =   120
      TabIndex        =   21
      Top             =   3405
      Width           =   765
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      X1              =   8
      X2              =   552
      Y1              =   256
      Y2              =   256
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      X1              =   8
      X2              =   552
      Y1              =   255
      Y2              =   255
   End
   Begin VB.Label labMisc 
      AutoSize        =   -1  'True
      Caption         =   "Texture 4:"
      Height          =   195
      Index           =   8
      Left            =   120
      TabIndex        =   19
      Top             =   2325
      Width           =   765
   End
   Begin VB.Label labMisc 
      AutoSize        =   -1  'True
      Caption         =   "Texture 6:"
      Height          =   195
      Index           =   7
      Left            =   120
      TabIndex        =   17
      Top             =   3045
      Width           =   765
   End
   Begin VB.Label labMisc 
      AutoSize        =   -1  'True
      Caption         =   "Texture 5:"
      Height          =   195
      Index           =   6
      Left            =   120
      TabIndex        =   16
      Top             =   2685
      Width           =   765
   End
   Begin VB.Label labMisc 
      AutoSize        =   -1  'True
      Caption         =   "Texture 3:"
      Height          =   195
      Index           =   5
      Left            =   120
      TabIndex        =   15
      Top             =   1965
      Width           =   765
   End
   Begin VB.Label labMisc 
      AutoSize        =   -1  'True
      Caption         =   "Texture 2:"
      Height          =   195
      Index           =   4
      Left            =   120
      TabIndex        =   14
      Top             =   1605
      Width           =   765
   End
   Begin VB.Label labMisc 
      AutoSize        =   -1  'True
      Caption         =   "Texture 1:"
      Height          =   195
      Index           =   3
      Left            =   120
      TabIndex        =   13
      Top             =   1245
      Width           =   765
   End
   Begin VB.Label labMisc 
      AutoSize        =   -1  'True
      Caption         =   "Technique:"
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   12
      Top             =   885
      Width           =   795
   End
   Begin VB.Label labMisc 
      AutoSize        =   -1  'True
      Caption         =   "Shader Name:"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   8
      Top             =   525
      Width           =   1020
   End
   Begin VB.Label labMisc 
      AutoSize        =   -1  'True
      Caption         =   "Transparency:"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   135
      Width           =   1050
   End
End
Attribute VB_Name = "frmMaterial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Me.cbbTransparency.AddItem "None"
    Me.cbbTransparency.AddItem "Alpha Blend"
    Me.cbbTransparency.AddItem "Alpha Test"
    
    UpdateGui
End Sub


'apply changes
Private Sub cmdApply_Click()
    On Error GoTo errhandler
    
    Dim i As Long
    Dim reloadmaps As Boolean
    
    With vmesh
        If Not .loadok Then Exit Sub
        
        With .geom(selgeom)
            With .lod(sellod)
                With .mat(selmat)
                    
                    'alphamode/shader/technique
                    .alphamode = Me.cbbTransparency.ListIndex
                    .fxfile = Me.txtShader.Text
                    .technique = Me.txtTechnique.Text
                    
                    'determine number of maps
                    Dim mapcount As Long
                    For i = 0 To Me.txtMap.UBound
                        If Len(Me.txtMap.Item(i).Text) > 0 Then
                            mapcount = mapcount + 1
                        End If
                    Next i
                    
                    'reallocate number of maps
                    .mapnum = mapcount
                    ReDim Preserve .map(0 To .mapnum - 1)
                    ReDim Preserve .texmapid(0 To .mapnum - 1)
                    ReDim Preserve .mapuvid(0 To .mapnum - 1)
                    
                    'set maps
                    For i = 0 To .mapnum - 1
                        If .map(i) <> Me.txtMap.Item(i).Text Then
                            .map(i) = Me.txtMap.Item(i).Text
                            
                            'for some reason, texture need to be reloaded twice or the old texture
                            ' is still used in the viewport, so we clear it here as workaround
                            .texmapid(i) = 0
                            .mapuvid(i) = 0
                            
                            reloadmaps = True
                        End If
                    Next i
                    
                    'clamp seltex in case we have the texture selected that we removed
                    If seltex > .mapnum - 1 Then
                        seltex = .mapnum - 1
                    End If
                    
                End With
                
                'build
                BuildShader .mat(selmat), vmesh.filename
                
            End With
        End With
    End With
    
    If reloadmaps Then
        LoadMeshTextures
    End If
    frmMain.FillTreeView
    SetStatus "info", "Done."
    
    Unload Me
    
    Exit Sub
errhandler:
    MsgBox "frmMaterial::cmdApply_Click" & vbLf & err.description, vbCritical
End Sub


'close dialog
Private Sub cmdCancel_Click()
     Unload Me
End Sub


'updates gui with data
Private Sub UpdateGui()
    On Error GoTo errhandler
    
    With vmesh
        If Not .loadok Then Exit Sub
        If selgeom < 0 Then Exit Sub
        If sellod < 0 Then Exit Sub
        If selmat < 0 Then Exit Sub
        
        With .geom(selgeom)
            With .lod(sellod)
                With .mat(selmat)
                    
                    Me.cbbTransparency.ListIndex = .alphamode
                    Me.txtShader.Text = .fxfile
                    Me.txtTechnique.Text = .technique
                    
                    Dim i As Long
                    For i = 0 To .mapnum - 1
                        Me.txtMap.Item(i).Text = .map(i)
                    Next i
                    
                End With
            End With
        End With
    End With
    
    Exit Sub
errhandler:
    MsgBox "frmMaterial::Update" & vbLf & err.description, vbCritical
End Sub


'browse
Private Sub cmdBrowse_Click(Index As Integer)

    Dim defpath As String
    Dim basepath As String
    
    'get current texture
    Dim fname As String
    fname = BF2GetTextureFilename(selgeom, sellod, selmat, Index)
    
    'determine base path
    If Not FileExist(fname) Then
        fname = vmesh.filename
        defpath = vmesh.filename
    Else
        fname = CleanFilePath(fname)
        
        Dim oldname As String
        oldname = CleanFilePath(Me.txtMap(Index).Text)
        
        If Len(oldname) > 0 Then
            
            'get base folder of old relative texture path
            Dim oldbase As String
            Dim i As Long
            For i = 1 To Len(oldname)
                Dim c As String
                c = Mid(oldname, i, 1)
                If c = "/" Then
                    If Len(oldbase) > 0 Then Exit For
                Else
                    oldbase = oldbase & c
                End If
            Next i
            If Len(oldbase) > 0 Then
                
                'ensure we don't mix up filenames with folders
                oldbase = "/" & oldbase & "/"
                
                'get part of path before oldbase
                Dim loc As Long
                loc = InStr(1, fname, oldbase)
                If loc > 0 Then
                    basepath = Left(fname, loc - 1)
                    If Right(basepath, 1) <> "/" Then
                        basepath = basepath & "/"
                    End If
                End If
            End If
            
        End If
        
        defpath = fname
    End If
    
    With Me.cdlFile
        .DialogTitle = "Browse"
        .Filter = "Direct Draw Surface (*.dds)|*.dds|" & _
                  "All Files (*.*)|*.*"
        .FilterIndex = 0
        .flags = cdlOFNExplorer Or cdlOFNFileMustExist
        .InitDir = Replace(defpath, "/", "\")
        .filename = GetFileName(fname)
        .CancelError = True
        On Error Resume Next
        .ShowOpen
        If Not err.Number = cdlCancel Then
            On Error GoTo 0
            
            'get new texture filename
            Dim newname As String
            newname = CleanFilePath(.filename)
            
            'construct relative path
            Dim relpath As String
            relpath = Replace(newname, basepath, "")
            
            'copy path to textbox
            Me.txtMap(Index).Text = relpath
        End If
    End With
End Sub

