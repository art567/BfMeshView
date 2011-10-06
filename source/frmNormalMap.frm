VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmNormalMap 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Convert Normal Map"
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8055
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmNormalMap.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   185
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   537
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog cdlFile 
      Left            =   2400
      Top             =   2280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame fraMisc 
      Caption         =   "Convert"
      Height          =   2055
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   7815
      Begin VB.ComboBox cbbMat 
         Height          =   315
         Left            =   6480
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   1170
         Width           =   1215
      End
      Begin VB.ComboBox cbbLod 
         Height          =   315
         Left            =   5160
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   1170
         Width           =   1215
      End
      Begin VB.ComboBox cbbGeom 
         Height          =   315
         Left            =   3840
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   1170
         Width           =   1215
      End
      Begin VB.TextBox txtFlatten 
         Height          =   285
         Left            =   2400
         TabIndex        =   13
         Text            =   "0"
         Top             =   1170
         Width           =   735
      End
      Begin VB.TextBox txtPadding 
         Height          =   285
         Left            =   960
         TabIndex        =   11
         Text            =   "5"
         Top             =   1170
         Width           =   615
      End
      Begin VB.CommandButton cmdOutput 
         Caption         =   "Browse..."
         Height          =   375
         Left            =   6600
         TabIndex        =   10
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox txtOutput 
         Height          =   285
         Left            =   960
         TabIndex        =   9
         Top             =   750
         Width           =   5535
      End
      Begin VB.CommandButton cmdInput 
         Caption         =   "Browse..."
         Height          =   375
         Left            =   6600
         TabIndex        =   5
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtInput 
         Height          =   285
         Left            =   960
         TabIndex        =   4
         Top             =   270
         Width           =   5535
      End
      Begin ComctlLib.ProgressBar pgbBar 
         Height          =   255
         Left            =   960
         TabIndex        =   6
         Top             =   1620
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   450
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.Label labMisc 
         AutoSize        =   -1  'True
         Caption         =   "Flatten:"
         Height          =   195
         Index           =   4
         Left            =   1680
         TabIndex        =   14
         Top             =   1200
         Width           =   570
      End
      Begin VB.Label labMisc 
         AutoSize        =   -1  'True
         Caption         =   "Padding:"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   12
         Top             =   1200
         Width           =   630
      End
      Begin VB.Label labMisc 
         AutoSize        =   -1  'True
         Caption         =   "Output:"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   8
         Top             =   780
         Width           =   570
      End
      Begin VB.Label labMisc 
         AutoSize        =   -1  'True
         Caption         =   "Progress:"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   1620
         Width           =   690
      End
      Begin VB.Label labMisc 
         AutoSize        =   -1  'True
         Caption         =   "Input:"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   300
         Width           =   450
      End
   End
   Begin VB.CommandButton cmdConvert 
      Caption         =   "Convert"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   2310
      Width           =   1095
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   6840
      TabIndex        =   0
      Top             =   2310
      Width           =   1095
   End
End
Attribute VB_Name = "frmNormalMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private nmap_folder As String

Private tgahead As tga_header

Private Type mat3
    m00 As Single
    m01 As Single
    m02 As Single
    m10 As Single
    m11 As Single
    m12 As Single
    m20 As Single
    m21 As Single
    m22 As Single
End Type

Private Type bgr
    b As Byte
    g As Byte
    r As Byte
End Type

Private Type bgra
    b As Byte
    g As Byte
    r As Byte
    a As Byte
End Type

'--- form ------------------------------------------------------------------------

Private Sub Form_Load()
    nmap_folder = current_folder
    Me.txtInput.Text = nmap_lastinput
    Me.txtOutput.Text = nmap_lastoutput
    Me.txtPadding.Text = nmap_padding
    Me.txtFlatten.Text = nmap_flatten
    
    With vmesh
        If .loadok Then
            Dim i As Long
            For i = 0 To .geomnum - 1
                Me.cbbGeom.AddItem "Geom " & i
            Next i
            Me.cbbGeom.ListIndex = 0
        End If
    End With
End Sub

Private Sub cbbGeom_Click()
    With vmesh
        If Not .loadok Then Exit Sub
        
        Me.cbbLod.Clear
        
        If .geomnum = 0 Then Exit Sub
        
        Dim i As Long
        For i = 0 To .geom(cbbGeom.ListIndex).lodnum - 1
            Me.cbbLod.AddItem "LOD " & i
        Next i
        Me.cbbLod.ListIndex = 0
    End With
End Sub

Private Sub cbbLod_Click()
    With vmesh
        If Not .loadok Then Exit Sub
        If .geomnum = 0 Then Exit Sub
        
        Me.cbbMat.Clear
        Dim i As Long
        For i = 0 To .geom(cbbGeom.ListIndex).lod(cbbLod.ListIndex).matnum - 1
            Me.cbbMat.AddItem "Mat " & i
        Next i
        Me.cbbMat.ListIndex = 0
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    nmap_lastinput = Me.txtInput.Text
    nmap_lastoutput = Me.txtOutput.Text
    nmap_padding = val(Me.txtPadding.Text)
    nmap_flatten = val(Me.txtFlatten.Text)
End Sub

Private Sub cmdInput_Click()
    With Me.cdlFile
        .DialogTitle = "Output"
        .Filter = "TGA (*.tga)|*.tga"
        .FilterIndex = 1
        .flags = cdlOFNExplorer Or cdlOFNFileMustExist
        .InitDir = nmap_folder
        .CancelError = True
        .filename = ""
        On Error Resume Next
        .ShowOpen
        If Not err.Number = cdlCancel Then
            On Error GoTo 0
            
            nmap_folder = .filename
            Me.txtInput.Text = .filename
        End If
    End With
End Sub

Private Sub cmdOutput_Click()
    With Me.cdlFile
        .DialogTitle = "Output"
        .Filter = "TGA (*.tga)|*.tga"
        .FilterIndex = 1
        .flags = cdlOFNOverwritePrompt Or cdlOFNNoReadOnlyReturn
        .InitDir = nmap_folder
        .CancelError = True
        .filename = ""
        On Error Resume Next
        .ShowSave
        If Not err.Number = cdlCancel Then
            On Error GoTo 0
            
            nmap_folder = .filename
            Me.txtOutput.Text = .filename
        End If
    End With
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdConvert_Click()
    If Not vmesh.loadok Then
        MsgBox "No model loaded.", vbExclamation
        Exit Sub
    End If
    
    cmdConvert.Enabled = False
    cmdClose.Enabled = False
    
    ConvertMap
    
    cmdConvert.Enabled = True
    cmdClose.Enabled = True
End Sub

'--- conversion -----------------------------------------------------------------


'dumps tangents to file
Public Sub ConvertMap()
    'On Error GoTo 0 errorhandler
    
Dim outputname As String
Dim inputname As String
    outputname = Me.txtOutput.Text
    inputname = Me.txtInput.Text
    
Dim flatten As Single
    flatten = val(Me.txtFlatten.Text)
    
    '--- open input normalmap ------------------------------------------------------------
    
    If Not FileExist(inputname) Then
        MsgBox "File " & Chr(34) & inputname & Chr(34) & " not found.", vbExclamation
        Exit Sub
    End If
    
    'open input file
    Dim ffin As Integer
    ffin = FreeFile
    Open inputname For Binary Access Read Lock Write As #ffin
    
    'read header
    Get #ffin, , tgahead
    If (tgahead.colortype <> 0) Or (tgahead.imagetype <> 2) Then
        MsgBox "Input file must be 24-bit RGB!", vbExclamation
        Close ffin
    End If
    
    'copy info
    Dim w As Long
    Dim h As Long
    Dim size As Long
    w = tgahead.width
    h = tgahead.height
    size = w * h
    
    'read data
    Dim data() As bgr
    ReDim data(0 To size - 1)
    Get #ffin, , data()
    
    'close input
    Close ffin
    
    '--- rasterize tangent space matrices -----------------------------------------------
    
    Dim geomid As Long
    Dim lodid As Long
    Dim matid As Long
    geomid = Me.cbbGeom.ListIndex
    lodid = Me.cbbLod.ListIndex
    matid = Me.cbbMat.ListIndex
    
    'allocate
    Dim mat() As mat3
    Dim flag() As Boolean
    Dim flip() As Boolean
    ReDim mat(0 To size - 1)
    ReDim flag(0 To size - 1)
    ReDim flip(0 To size - 1)
    
    'clear stuff
    Dim i As Long
    For i = 0 To size - 1
        SetIdentity mat(i)
        flag(i) = False
    Next i
    
    'rasterize
    Dim stride As Long
    
    Dim vstart As Long
    Dim istart As Long
    Dim facenum As Long
    
    Dim normoff As Long
    Dim texcoff As Long
    Dim tangoff As Long
    
    Dim vi As Long
    Dim v1 As Long
    Dim v2 As Long
    Dim v3 As Long
    
    Dim t1 As float2
    Dim t2 As float2
    Dim t3 As float2
    
    Dim n As float3
    Dim n1 As float3  'normal vector
    Dim n2 As float3  'normal vector
    Dim n3 As float3  'normal vector
    
    Dim tv As float3
    Dim tv1 As float3 'tangent vector
    Dim tv2 As float3 'tangent vector
    Dim tv3 As float3 'tangent vector
    
    Dim bt As float3
    Dim bt1 As float3 'bitangent vector
    Dim bt2 As float3 'bitangent vector
    Dim bt3 As float3 'bitangent vector
    
    Dim minx As Single
    Dim miny As Single
    Dim maxx As Single
    Dim maxy As Single
    
    Dim x As Long
    Dim y As Long
    Dim j As Long
    
    Dim p As float2
    Dim sx As Single
    Dim sy As Single
    Dim ox As Single
    Dim oy As Single
    
    'compute scale
    sx = 1 / w
    sy = 1 / h
    
    'compute offset
    ox = sx / 2
    oy = sy / 2
    
    With vmesh
        
        stride = .vertstride / 4
        
        normoff = BF2MeshGetNormOffset
        texcoff = BF2MeshGetTexcOffset(0)
        tangoff = BF2MeshGetTangOffset
        
        vstart = .geom(geomid).lod(lodid).mat(matid).vstart
        istart = .geom(geomid).lod(lodid).mat(matid).istart
        
        facenum = .geom(geomid).lod(lodid).mat(matid).inum / 3
        
        For i = 0 To facenum - 1
            
            Me.pgbBar.value = (i / facenum) * 100
            
            'get face vert indices
            v1 = (vstart + .Index(istart + (i * 3) + 0)) * stride
            v2 = (vstart + .Index(istart + (i * 3) + 1)) * stride
            v3 = (vstart + .Index(istart + (i * 3) + 2)) * stride
            
            'get face texcoords
            n1.x = .vert(v1 + normoff + 0)
            n1.y = .vert(v1 + normoff + 1)
            n1.z = .vert(v1 + normoff + 2)
            t1.x = .vert(v1 + texcoff + 0)
            t1.y = .vert(v1 + texcoff + 1)
            tv1.x = .vert(v1 + tangoff + 0)
            tv1.y = .vert(v1 + tangoff + 1)
            tv1.z = .vert(v1 + tangoff + 2)
            bt1 = Normalize(CrossProduct(n1, tv1))
            
            n2.x = .vert(v2 + normoff + 0)
            n2.y = .vert(v2 + normoff + 1)
            n2.z = .vert(v2 + normoff + 2)
            t2.x = .vert(v2 + texcoff + 0)
            t2.y = .vert(v2 + texcoff + 1)
            tv2.x = .vert(v2 + tangoff + 0)
            tv2.y = .vert(v2 + tangoff + 1)
            tv2.z = .vert(v2 + tangoff + 2)
            bt2 = Normalize(CrossProduct(n2, tv2))
            
            n3.x = .vert(v3 + normoff + 0)
            n3.y = .vert(v3 + normoff + 1)
            n3.z = .vert(v3 + normoff + 2)
            t3.x = .vert(v3 + texcoff + 0)
            t3.y = .vert(v3 + texcoff + 1)
            tv3.x = .vert(v3 + tangoff + 0)
            tv3.y = .vert(v3 + tangoff + 1)
            tv3.z = .vert(v3 + tangoff + 2)
            bt3 = Normalize(CrossProduct(n3, tv3))
            
            'compute triangle rect bounds
            minx = min(min(t1.x, t2.x), t3.x) * (w - 1)
            miny = min(min(t1.y, t2.y), t3.y) * (h - 1)
            maxx = max(max(t1.x, t2.x), t3.x) * (w - 1)
            maxy = max(max(t1.y, t2.y), t3.y) * (h - 1)
            minx = Clamp(minx, 0, w - 1)
            miny = Clamp(miny, 0, h - 1)
            maxx = Clamp(maxx, 0, w - 1)
            maxy = Clamp(maxy, 0, h - 1)
            
            'rasterize
            For x = minx To maxx
                For y = miny To maxy
                    
                    'compute UV position (texel center)
                    p.x = (x * sx) + ox
                    p.y = (y * sy) + oy
                    
                    Dim insideCW As Boolean
                    Dim insideCCW As Boolean
                    insideCW = TriangleTestCW(t1, t2, t3, p)
                    If Not insideCW Then
                        insideCCW = TriangleTestCW(t3, t2, t1, p)
                    End If
                    If insideCW Or insideCCW Then
                        
                        'compute sample index
                        j = x + (((w - 1) - y) * w)
                        
                        'interpolate
                        n = Normalize(TexelToPoint(n1, n2, n3, t1, t2, t3, p))
                        tv = Normalize(TexelToPoint(tv1, tv2, tv3, t1, t2, t3, p))
                        bt = Normalize(TexelToPoint(bt1, bt2, bt3, t1, t2, t3, p))
                        
                        'set matrix
                        mat(j).m00 = -n.x
                        mat(j).m01 = n.y
                        mat(j).m02 = n.z
                        
                        mat(j).m10 = -tv.x
                        mat(j).m11 = tv.y
                        mat(j).m12 = tv.z
                        
                        mat(j).m20 = -bt.x
                        mat(j).m21 = bt.y
                        mat(j).m22 = bt.z
                        
                        flag(j) = True
                        flip(j) = insideCCW
                        
                    End If
                    
                Next y
            Next x
            
        Next i
        
    End With
    
    '--- convert --------------------------------------------------
    
    'On Error Resume Next
    
    Dim pixel() As bgra
    ReDim pixel(0 To size - 1)
    
    'write data
    Dim vec As float3
    Dim c As bgra
    For i = 0 To size - 1
        Me.pgbBar.value = (i / size) * 100
        
        If flag(i) Then
            
            If data(i).r = 127 And data(i).g = 127 And data(i).b = 127 Then
                'ignore gray
                c.r = 127
                c.g = 127
                c.b = 255
                c.a = 127
            Else
                
                'get original object space normal
                vec.x = ((CSng(data(i).r) / 255) * 2) - 1
                vec.y = ((CSng(data(i).g) / 255) * 2) - 1
                vec.z = ((CSng(data(i).b) / 255) * 2) - 1
                'vec = Normalize(vec)
                
                'rotate object space normal
                vec = RotateVectorInverse(mat(i), vec)
                
                'swap channels
                Dim tmp As float3
                tmp = vec
                vec.x = tmp.y
                vec.y = -tmp.z
                vec.z = tmp.x
                
                'apply flattening
                vec.z = vec.z + flatten
                
                'normalize result
                vec = Normalize(vec)
                
                'flip if UVs inverted
                'If flip(i) Then
                '    'vec.x = -vec.x
                '    'vec.y = -vec.y
                'End If
                
                'store the resulting vector
                c.r = ((vec.x + 1) / 2) * 255
                c.g = ((vec.y + 1) / 2) * 255
                c.b = ((vec.z + 1) / 2) * 255
                c.a = 255
                
            End If
            
        Else
            c.r = 127
            c.g = 127
            c.b = 255
            c.a = 0
        End If
        
        pixel(i) = c
    Next i
    
    '--- apply padding ------------------------------------------
    
    GenPadding w, h, pixel()
    
    '--- write output -------------------------------------------
    
    'create file
    Dim ffout As Integer
    ffout = FreeFile
    Open outputname For Binary As #ffout
    
    'write header
    tgahead.flip = 0
    tgahead.bits = 32
    Put #ffout, , tgahead
    Put #ffout, , pixel()
    
    'close file
    Close ffout
    
    'clean up
    Erase data()
    Erase mat()
    
    'done
    'MsgBox "Done.", vbInformation
    Me.pgbBar.value = 0
    
    Exit Sub
errorhandler:
    MsgBox "ConvertMap" & vbLf & err.description, vbCritical
End Sub


'sets matrix identity
Private Sub SetIdentity(ByRef mat As mat3)
    mat.m00 = 1:    mat.m01 = 0:    mat.m02 = 0
    mat.m10 = 0:    mat.m11 = 1:    mat.m12 = 0
    mat.m20 = 0:    mat.m21 = 0:    mat.m22 = 1
End Sub


'rotates vector by matrix
Private Function RotateVector(ByRef mat As mat3, ByRef v As float3) As float3
Dim r As float3
    r.x = v.x * mat.m00 + v.y * mat.m10 + v.z * mat.m20
    r.y = v.x * mat.m01 + v.y * mat.m11 + v.z * mat.m21
    r.z = v.x * mat.m02 + v.y * mat.m12 + v.z * mat.m22
    RotateVector = r
End Function

'rotates vector by matrix
Private Function RotateVectorInverse(ByRef mat As mat3, ByRef v As float3) As float3
Dim r As float3
    r.x = (v.x * mat.m00) + (v.y * mat.m01) + (v.z * mat.m02)
    r.y = (v.x * mat.m10) + (v.y * mat.m11) + (v.z * mat.m12)
    r.z = (v.x * mat.m20) + (v.y * mat.m21) + (v.z * mat.m22)
    RotateVectorInverse = r
End Function

'returns the determinant
Private Function Determinant(ByRef mat As mat3) As Single
    Determinant = mat.m00 * mat.m11 * mat.m22 + _
                  mat.m01 * mat.m12 * mat.m20 + _
                  mat.m02 * mat.m20 * mat.m21 - _
                  mat.m00 * mat.m12 * mat.m21 - _
                  mat.m01 * mat.m10 * mat.m22 - _
                  mat.m02 * mat.m11 * mat.m20
End Function

'inverts matrix
'Private Sub InvertMatrix(ByRef mat As mat3)
'Dim tmp As mat3
'
'    '1 / Determinant(mat)
'
'End Sub

'computes pixel index from coordinates
Private Function pPos(ByVal x As Long, ByVal y As Long, ByRef w As Long) As Long
    pPos = x + (w * y)
End Function


'applies padding to texture
Private Sub GenPadding(ByVal w As Long, ByVal h As Long, ByRef data() As bgra)
Dim i As Long
Dim x As Long
Dim y As Long
Dim n As Long
Dim cr As Long
Dim cg As Long
Dim cb As Long
Dim a As Long
Dim p As Long
    On Error GoTo errhandler
    
    Dim padding As Long
    padding = val(Me.txtPadding.Text)
    
    For i = 1 To padding
        For x = 0 To w - 1
            For y = 0 To h - 1
                p = x + (y * w)
                If data(p).a = 0 Then
                    
                    cr = 0
                    cg = 0
                    cb = 0
                    n = 0
                    
                    'north
                    If y > 0 Then
                        p = x + (w * (y - 1))
                        If data(p).a = 255 Then
                            cr = cr + data(p).r
                            cg = cg + data(p).g
                            cb = cb + data(p).b
                            n = n + 1
                        End If
                    End If
                    
                    'south
                    If y < h - 1 Then
                        p = x + (w * (y + 1))
                        If data(p).a = 255 Then
                            cr = cr + data(p).r
                            cg = cg + data(p).g
                            cb = cb + data(p).b
                            n = n + 1
                        End If
                    End If
                    
                    'west
                    If x > 0 Then
                        p = (x - 1) + (y * w)
                        If data(p).a = 255 Then
                            cr = cr + data(p).r
                            cg = cg + data(p).g
                            cb = cb + data(p).b
                            n = n + 1
                        End If
                    End If
                    
                    'east
                    If x < w - 1 Then
                        p = (x + 1) + (y * w)
                        If data(p).a = 255 Then
                            cr = cr + data(p).r
                            cg = cg + data(p).g
                            cb = cb + data(p).b
                            n = n + 1
                        End If
                    End If
                    
                    'set pixel
                    If n > 0 Then
                        p = x + (y * w)
                        data(p).r = cr / n
                        data(p).g = cg / n
                        data(p).b = cb / n
                        data(p).a = 127
                    End If
                    
                End If
            Next y
        Next x
        
        'update alpha
        For x = 0 To w - 1
            For y = 0 To h - 1
                p = x + (y * w)
                If data(p).a = 127 Then
                    data(p).a = 255
                End If
            Next y
        Next x
        
    Next i
    
    Exit Sub
errhandler:
    MsgBox err.description
End Sub

