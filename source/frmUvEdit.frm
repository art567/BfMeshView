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
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Tools"
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7575
      Begin VB.CommandButton cmdScale 
         Caption         =   "Scale"
         Height          =   315
         Left            =   5760
         TabIndex        =   6
         Top             =   60
         Width           =   855
      End
      Begin VB.CommandButton cmdRotate 
         Caption         =   "Rotate"
         Height          =   315
         Left            =   4800
         TabIndex        =   5
         Top             =   60
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
         Left            =   6720
         TabIndex        =   2
         Top             =   60
         Width           =   855
      End
      Begin MSComctlLib.Toolbar tlbTools 
         Height          =   330
         Left            =   2040
         TabIndex        =   4
         Top             =   45
         Width           =   2685
         _ExtentX        =   4736
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "imlTools"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   9
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Select UVs"
               ImageIndex      =   1
               Style           =   3
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "select"
               Object.ToolTipText     =   "Select UVs"
               ImageIndex      =   1
               Style           =   2
               Value           =   1
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "move"
               Object.ToolTipText     =   "Move UVs"
               ImageIndex      =   2
               Style           =   2
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "scale"
               Object.ToolTipText     =   "Scale UVs"
               ImageIndex      =   3
               Style           =   2
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "uvxneg"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "uvxpos"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "uvyneg"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
      Left            =   1080
      ScaleHeight     =   300
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   300
      TabIndex        =   0
      Top             =   1200
      Width           =   4500
      Begin VB.Shape shpSel 
         BorderColor     =   &H00FFFFFF&
         BorderStyle     =   3  'Dot
         Height          =   735
         Left            =   360
         Top             =   360
         Visible         =   0   'False
         Width           =   1575
      End
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
Private dragx As Long 'mouse coordinates at start of drag
Private dragy As Long 'mouse coordinates at start of drag
Private dsx As Single
Private dsy As Single
Private toolmode As e_toolmode
Private uvchan As Long

Private keyctrl As Boolean
Private keyalt As Boolean


Private Sub Form_Load()
    'MsgBox Me.Caption
    
    uveditor_isloaded = True
    
    offx = 50
    offy = 50
    zoom = 200
    uvchan = 0
    toolmode = 0
    
    With vmesh
        If .loadok Then
            Dim i As Long
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
            Me.cbbChannel.ListIndex = 0
        End If
    End With
    
    SetTopMostWindow Me.hWnd, True
    
    Center Me
End Sub

Private Sub cbbChannel_Click()
    uvchan = Me.cbbChannel.ListIndex
    If uvchan < 0 Then uvchan = 0
    If uvchan > 4 Then uvchan = 4
    
    ClearVertSelection
    
    picMain_Paint
    frmMain.picMain_Paint
    
    'Me.picMain.SetFocus
End Sub

Private Sub Form_Resize()
    If Not Me.WindowState = vbMinimized Then
        
        If Me.width < 200 * 15 Then Me.width = 200 * 15
        If Me.height < 100 * 15 Then Me.height = 100 * 15
        
        Me.picMain.Move 3, 30, Me.ScaleWidth - 6, Me.ScaleHeight - 30 - 3
        
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not app_exit Then
        Cancel = True
        Me.Hide
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
        
        If Button = vbLeftButton Then
            Select Case toolmode
            Case tool_select
                
                Dim minx As Long
                Dim miny As Long
                Dim maxx As Long
                Dim maxy As Long
                
                minx = min(dragx, X)
                miny = min(dragy, Y)
                maxx = max(dragx, X)
                maxy = max(dragy, Y)
                
                Me.shpSel.Move minx, miny, maxx - minx, maxy - miny
                Me.shpSel.Visible = True
                
            Case tool_move
                MoveVerts (X - mousex) / zoom, (Y - mousey) / zoom
                
            Case tool_scale
                ScaleVerts (X - mousex) / zoom, (Y - mousey) / zoom
                
            End Select
        End If
        
        If Button = vbRightButton Then
            zoom = zoom - (Y - mousey)
            If zoom < 0 Then zoom = 0
            If zoom < 0 Then zoom = 0
        End If
        
        If Button = vbMiddleButton Then
            offx = offx + (X - mousex)
            offy = offy + (Y - mousey)
        End If
        
        picMain_Paint
        frmMain.picMain_Paint
    End If
    mousex = X
    mousey = Y
End Sub


Private Sub picMain_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    'selection
    If Button = vbLeftButton Then
        If toolmode = tool_select Then
            Dim minx As Long
            Dim miny As Long
            Dim maxx As Long
            Dim maxy As Long
            
            minx = min(dragx, X)
            miny = min(dragy, Y)
            maxx = max(dragx, X)
            maxy = max(dragy, Y)
            
            SelVerts (minx - offx) / zoom, (miny - offy) / zoom, _
                     (maxx - offx) / zoom, (maxy - offy) / zoom
        End If
    End If
    
    Me.shpSel.Visible = False
    mousedown = False
    picMain_Paint
    frmMain.picMain_Paint
    
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
        
        'uvoffset = 7 + (uvchan * 2)
        uvoffset = BF2MeshGetTexcOffset(uvchan)
        stride = .vertstride / 4
        
        With .geom(selgeom).lod(sellod)
            
            Dim matmin As Long
            Dim matmax As Long
            If selmat < 0 Then
                matmin = 0
                matmax = .matnum - 1
            Else
                matmin = selmat
                matmax = selmat
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
    If Not Me.Visible Then Exit Sub
    
    On Error GoTo errhandler:
    
Dim i As Long
    
    'LockWindowUpdate Me.picMain.hWnd
    picMain.Cls
    
    'draw scale
    dsx = zoom
    dsy = zoom
    
    With vmesh
        If Not .loadok Then Exit Sub
        If selgeom < 0 Then Exit Sub
        If sellod < 0 Then Exit Sub
        
        Dim stride As Long
        Dim uvoffset As Long
        
        'uvoffset = 7 + (uvchan * 2)
        uvoffset = BF2MeshGetTexcOffset(uvchan)
        stride = .vertstride / 4
        
        'draw border
        picMain.ForeColor = RGB(0, 0, 0)
        DrawLine 0, 0, 1, 0
        DrawLine 0, 0, 0, 1
        DrawLine 1, 1, 1, 0
        DrawLine 1, 1, 0, 1
                
        'draw triangles
        picMain.ForeColor = RGB(50, 200, 50)
        With .geom(selgeom).lod(sellod)
            
            Dim matmin As Long
            Dim matmax As Long
            If selmat < 0 Then
                matmin = 0
                matmax = .matnum - 1
            Else
                matmin = selmat
                matmax = selmat
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
                        
                        DrawTri vmesh.vert(f1 + 0), vmesh.vert(f1 + 1), _
                                vmesh.vert(f2 + 0), vmesh.vert(f2 + 1), _
                                vmesh.vert(f3 + 0), vmesh.vert(f3 + 1)
                    Next i
                End With
            Next m
        End With
        
        'draw vertices
        SetVertFlags
        For i = 0 To .vertnum - 1
            If .vertflag(i) Then
            
                Dim X As Single
                Dim Y As Single
                X = .vert(i * stride + uvoffset + 0)
                Y = .vert(i * stride + uvoffset + 1)
                
                If .vertsel(i) Then
                    picMain.ForeColor = RGB(255, 0, 0)
                Else
                    picMain.ForeColor = RGB(255, 255, 255)
                End If
                
                DrawVert X, Y
                
            End If
        Next i
        
    End With
    'LockWindowUpdate 0
    
    Exit Sub
errhandler:
    Me.Caption = "picMain_Paint Error: " & err.description
    On Error GoTo 0
End Sub


'draws line between two points
Private Sub DrawLine(ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single)
    picMain.Line (x1 * dsx + offx, y1 * dsy + offy)-(x2 * dsx + offx, y2 * dsy + offy)
End Sub


'draws triangle
Private Sub DrawTri(ByVal v1x As Single, ByVal v1y As Single, _
                    ByVal v2x As Single, ByVal v2y As Single, _
                    ByVal v3x As Single, ByVal v3y As Single)
    picMain.Line (v1x * dsx + offx, v1y * dsy + offy)-(v2x * dsx + offx, v2y * dsy + offy)
    picMain.Line (v2x * dsx + offx, v2y * dsy + offy)-(v3x * dsx + offx, v3y * dsy + offy)
    picMain.Line (v3x * dsx + offx, v3y * dsy + offy)-(v1x * dsx + offx, v1y * dsy + offy)
End Sub


'draws vertex
Private Sub DrawVert(ByVal X As Single, ByVal Y As Single)
    'picMain.PSet (x * zoomscale + offx, y * zoomscale + offx), picMain.ForeColor
    picMain.Circle (X * dsx + offx, Y * dsy + offy), 1
    
    picMain.Circle (X * dsx + offx + 1, Y * dsy + offy + 0), 1
    picMain.Circle (X * dsx + offx + 0, Y * dsy + offy + 1), 1
    picMain.Circle (X * dsx + offx + 1, Y * dsy + offy + 1), 1
End Sub

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
