VERSION 5.00
Begin VB.Form frmSkin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Skin"
   ClientHeight    =   6855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3735
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSkin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   457
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   249
   Begin VB.CheckBox chkEditMode 
      Caption         =   "Select"
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6240
      Width           =   3495
   End
   Begin VB.CommandButton cmdAuto 
      Caption         =   "Auto-Assign"
      Height          =   375
      Left            =   1920
      TabIndex        =   14
      Top             =   5280
      Width           =   1695
   End
   Begin VB.CommandButton Command6 
      Caption         =   "0.75"
      Height          =   375
      Left            =   840
      TabIndex        =   13
      Top             =   4200
      Width           =   615
   End
   Begin VB.CommandButton Command5 
      Caption         =   "0.25"
      Height          =   375
      Left            =   2280
      TabIndex        =   12
      Top             =   4200
      Width           =   615
   End
   Begin VB.CommandButton Command4 
      Caption         =   "0.0"
      Height          =   375
      Left            =   3000
      TabIndex        =   11
      Top             =   4200
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      Caption         =   "0.5"
      Height          =   375
      Left            =   1560
      TabIndex        =   10
      Top             =   4200
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "1.0"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   4200
      Width           =   615
   End
   Begin VB.CommandButton cmdMirrorBlend 
      Caption         =   "Mirror Blend"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   5280
      Width           =   1695
   End
   Begin VB.CommandButton cmdMirrorToRight 
      Caption         =   "Mirror To Right"
      Height          =   375
      Left            =   1920
      TabIndex        =   7
      Top             =   4800
      Width           =   1695
   End
   Begin VB.CommandButton cmdMirrorToLeft 
      Caption         =   "Mirror To Left"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   4800
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Load..."
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   5760
      Width           =   1695
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save..."
      Height          =   375
      Left            =   1920
      TabIndex        =   4
      Top             =   5760
      Width           =   1695
   End
   Begin VB.PictureBox picWeight 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   300
      Left            =   120
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   233
      TabIndex        =   3
      Top             =   3780
      Width           =   3495
   End
   Begin VB.ComboBox cbbBone2 
      Height          =   315
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   3360
      Width           =   1695
   End
   Begin VB.ComboBox cbbBone1 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   3360
      Width           =   1695
   End
   Begin VB.ListBox lstBones 
      Height          =   2985
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3495
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      Index           =   3
      X1              =   8
      X2              =   240
      Y1              =   312
      Y2              =   312
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      Index           =   2
      X1              =   8
      X2              =   240
      Y1              =   313
      Y2              =   313
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      Index           =   1
      X1              =   8
      X2              =   240
      Y1              =   215
      Y2              =   215
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      Index           =   0
      X1              =   8
      X2              =   240
      Y1              =   216
      Y2              =   216
   End
End
Attribute VB_Name = "frmSkin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private skin_vert As Long
Private skin_weight As Single

Private Sub Form_Load()
    Me.Move 10 * 30, 50 * 30
    SetTopMostWindow Me.hWnd, True
    FillBoneList
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not app_exit Then
        Cancel = True
        Me.Hide
        Me.chkEditMode.value = vbUnchecked
    End If
End Sub

Private Sub picWeight_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picWeight_MouseMove Button, Shift, X, Y
End Sub

Private Sub picWeight_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button Then
        skin_weight = (X - 1) / (picWeight.width - 2)
        
        If skin_weight < 0 Then skin_weight = 0
        If skin_weight > 1 Then skin_weight = 1
        
        ChangeWeight
        
        picWeight_Paint
        frmMain.picMain_Paint
    End If
End Sub

Private Sub picWeight_Paint()
    Dim w As Long
    Dim h As Long
    w = picWeight.width
    h = picWeight.height
    
    'background
    picWeight.FillColor = RGB(255, 255, 255)
    picWeight.ForeColor = RGB(127, 127, 127)
    DrawRect picWeight.hDC, 0, 0, w, h
    
    'weight
    Dim wt As Long
    wt = (w - 2) * skin_weight
    If wt > 0 Then
        picWeight.FillColor = RGB(200, 200, 200)
        picWeight.ForeColor = RGB(200, 200, 200)
        DrawRect picWeight.hDC, 1, 1, wt + 1, h - 1
    End If
    
    'slider
    picWeight.FillColor = RGB(200, 63, 63)
    picWeight.ForeColor = RGB(200, 63, 63)
    DrawRect picWeight.hDC, wt + 1 - 2, 0, wt + 1 + 2, h
End Sub

Private Sub FillBoneList()
    Me.lstBones.Clear
    Me.cbbBone1.Clear
    Me.cbbBone2.Clear
    'Me.cbbBone1.AddItem "NONE"
    'Me.cbbBone2.AddItem "NONE"
    With bf2ske
        If Not .loaded Then Exit Sub
        
        Dim i As Long
        For i = 0 To .nodenum - 1
            
            Dim ident As Long
            ident = 0
            
            Dim p As Long
            p = .node(i).parent
            While p > -1
                p = .node(p).parent
                ident = ident + 1
            Wend
            
            Me.lstBones.AddItem (String(ident, " ") & .node(i).name)
            
            'Me.cbbBone1.AddItem .node(i).name
            'Me.cbbBone2.AddItem .node(i).name
        Next i
    End With
End Sub


Private Sub SelectSkinVert(ByRef v As Long)
    
    'fill bone lists
    Me.cbbBone1.Clear
    Me.cbbBone2.Clear
    With vmesh.geom(selgeom).lod(sellod).rig(vmesh.vertinfo(v).mat)
        Dim i As Long
        For i = 0 To .bonenum - 1
            Me.cbbBone1.AddItem bf2ske.node(.bone(i).id).name
            Me.cbbBone2.AddItem bf2ske.node(.bone(i).id).name
        Next i
    End With
    
    Dim vw As bf2skinweight
    GetSkinVertWeight v, vw
    
    Dim sb1 As Long
    Dim sb2 As Long
    'sb1 = .geom(vmesh.vertinfo(i).geom).lod(vmesh.vertinfo(i).lod).rig(vmesh.vertinfo(i).mat).bone(vw.b1).id
    'sb2 = .geom(vmesh.vertinfo(i).geom).lod(vmesh.vertinfo(i).lod).rig(vmesh.vertinfo(i).mat).bone(vw.b2).id
    
    Me.cbbBone1.ListIndex = vw.b1
    Me.cbbBone2.ListIndex = vw.b2
    
    skin_weight = vw.w
    
    'Me.Caption = "b1:" & vw.b1 & "|b2:" & vw.b2 & "|w" & vw.w
End Sub

Private Sub chkEditMode_Click()
    If chkEditMode.value = vbChecked Then
        toolmode = 1
    Else
        toolmode = 0
    End If
    
    'toggle stuff
    view_verts = (toolmode = 1)
    view_edges = (toolmode = 1)
    
    'redraw
    frmMain.picMain_Paint
End Sub


'updates after vertex selection
Public Sub SelectionChanged()
    On Error GoTo errhandler
    With vmesh
        If Not .loadok Then Exit Sub
        
        Dim i As Long
        For i = 0 To .vertnum - 1
            If .vertflag(i) Then
                If .vertsel(i) Then
                    
                    SelectSkinVert i
                    
                    Exit For
                End If
            End If
        Next i
        
        picWeight_Paint
    End With
    
    Exit Sub
errhandler:
    MsgBox "SelectionChanged" & vbLf & err.description, vbCritical
    'Me.Caption = err.description & " " & Rnd()
End Sub



Private Sub ChangeWeight()
    With vmesh
        If Not .loadok Then Exit Sub
        
        Dim vw As bf2skinweight
        
        Dim i As Long
        For i = 0 To .vertnum - 1
            If .vertflag(i) Then
                If .vertsel(i) Then
                    
                    GetSkinVertWeight i, vw
                    
                    vw.b1 = Me.cbbBone1.ListIndex
                    vw.b2 = Me.cbbBone2.ListIndex
                    vw.w = skin_weight
                    
                    SetSkinVertWeight i, vw
                    
                End If
            End If
        Next i
        
    End With
End Sub
