VERSION 5.00
Begin VB.Form frmTransform 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Vertex Transform"
   ClientHeight    =   1575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3375
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTransform.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   105
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtBoneID 
      Height          =   285
      Left            =   960
      TabIndex        =   0
      Text            =   "0"
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "Select"
      Height          =   375
      Left            =   2160
      TabIndex        =   4
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "Move"
      Height          =   375
      Left            =   2160
      TabIndex        =   5
      Top             =   1080
      Width           =   1095
   End
   Begin VB.TextBox txtOffset 
      BackColor       =   &H00FFC0C0&
      Height          =   285
      Index           =   2
      Left            =   960
      TabIndex        =   3
      Text            =   "0.0"
      Top             =   1200
      Width           =   975
   End
   Begin VB.TextBox txtOffset 
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Index           =   1
      Left            =   960
      TabIndex        =   2
      Text            =   "0.0"
      Top             =   840
      Width           =   975
   End
   Begin VB.TextBox txtOffset 
      BackColor       =   &H00C0C0FF&
      Height          =   285
      Index           =   0
      Left            =   960
      TabIndex        =   1
      Text            =   "0.0"
      Top             =   480
      Width           =   975
   End
   Begin VB.Label labMisc 
      AutoSize        =   -1  'True
      Caption         =   "Offset Z:"
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   9
      Top             =   1200
      Width           =   660
   End
   Begin VB.Label labMisc 
      AutoSize        =   -1  'True
      Caption         =   "Offset Y:"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   8
      Top             =   840
      Width           =   660
   End
   Begin VB.Label labMisc 
      AutoSize        =   -1  'True
      Caption         =   "Bone ID:"
      Height          =   195
      Index           =   3
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   630
   End
   Begin VB.Label labMisc 
      AutoSize        =   -1  'True
      Caption         =   "Offset X:"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   660
   End
End
Attribute VB_Name = "frmTransform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    SetTopMostWindow Me.hWnd, True
    
    cmdSelect_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMain.picMain_Paint
End Sub

Private Sub cmdSelect_Click()
    UpdateSelection
    frmMain.picMain_Paint
End Sub

Private Sub cmdMove_Click()
    UpdateSelection
    MoveVerts
    frmMain.picMain_Paint
End Sub

'------------------------------------------------------------------------------------------------------------

'sets the vertex flags of the currently selected geom+lod
Public Sub SetVertFlags2(ByRef geomid As Long, ByRef lodid As Long)
    On Error GoTo errhandler
    
    Dim i As Long
    With vmesh
        If Not .loadok Then Exit Sub
        If geomid < 0 Then Exit Sub
        If lodid < 0 Then Exit Sub
        
        'clear vert flags
        For i = 0 To .vertnum - 1
            .vertflag(i) = 0
        Next i
        
        '...
        Dim stride As Long
        stride = .vertstride / 4
        With .geom(geomid).lod(lodid)
            Dim m As Long
            For m = 0 To .matnum - 1
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
    MsgBox "SetVertFlags2" & vbLf & err.description, vbCritical
    On Error GoTo 0
End Sub


'sets vertex flags
Private Sub UpdateSelection()
    SetVertFlags2 selgeom, sellod
    
Dim i As Long
Dim stride As Long
Dim attriboff As Long
Dim vw As bf2vw

Dim ox As Single
Dim oy As Single
Dim oz As Single
    ox = val(Me.txtOffset(0).Text)
    oy = val(Me.txtOffset(1).Text)
    oz = val(Me.txtOffset(2).Text)
    
    Dim selid As Long
    selid = val(Me.txtBoneID.Text)
    
    With vmesh
        
        stride = .vertstride / 4
        
        attriboff = BF2MeshGetWeightOffset()
        
        For i = 0 To .vertnum - 1
            If .vertflag(i) > 0 Then
                
                'get block
                CopyMem VarPtr(vw), VarPtr(.vert(i * stride + attriboff)), 4
                
                If vw.b1 = selid Then
                    .vertsel(i) = 1
                Else
                    .vertsel(i) = 0
                End If
                
            End If
        Next i
        
    End With
End Sub


'move vertices by bone index
Private Sub MoveVerts()
    
Dim ox As Single
Dim oy As Single
Dim oz As Single
    ox = val(Me.txtOffset(0).Text)
    oy = val(Me.txtOffset(1).Text)
    oz = val(Me.txtOffset(2).Text)
    
    With vmesh
        
        Dim stride As Long
        stride = .vertstride / 4
        
        Dim i As Long
        For i = 0 To .vertnum - 1
            If .vertflag(i) > 0 Then
                If .vertsel(i) > 0 Then
                    
                    Dim v1 As Long
                    Dim v2 As Long
                    Dim v3 As Long
                    v1 = i * stride + 0
                    v2 = i * stride + 1
                    v3 = i * stride + 2
                    
                    'move vert
                    .vert(v1) = .vert(v1) + ox
                    .vert(v2) = .vert(v2) + oy
                    .vert(v3) = .vert(v3) + oz
                    
                End If
            End If
        Next i
        
    End With
End Sub

