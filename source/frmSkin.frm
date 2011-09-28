VERSION 5.00
Begin VB.Form frmSkin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Skin"
   ClientHeight    =   6255
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
   ScaleHeight     =   417
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   249
   Begin VB.CommandButton cmdAuto 
      Caption         =   "Auto-Assign"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   14
      Top             =   5280
      Width           =   1695
   End
   Begin VB.CommandButton Command6 
      Caption         =   "0.75"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   13
      Top             =   4200
      Width           =   615
   End
   Begin VB.CommandButton Command5 
      Caption         =   "0.25"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   12
      Top             =   4200
      Width           =   615
   End
   Begin VB.CommandButton Command4 
      Caption         =   "0.0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   11
      Top             =   4200
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      Caption         =   "0.5"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   10
      Top             =   4200
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "1.0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   4200
      Width           =   615
   End
   Begin VB.CommandButton cmdMirrorBlend 
      Caption         =   "Mirror Blend"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   5280
      Width           =   1695
   End
   Begin VB.CommandButton cmdMirrorToRight 
      Caption         =   "Mirror To Right"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   7
      Top             =   4800
      Width           =   1695
   End
   Begin VB.CommandButton cmdMirrorToLeft 
      Caption         =   "Mirror To Left"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   4800
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Load..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   5760
      Width           =   1695
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   3360
      Width           =   1695
   End
   Begin VB.ComboBox cbbBone1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   3360
      Width           =   1695
   End
   Begin VB.ListBox lstBones 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
    SetTopMostWindow Me.hWnd, True
    FillBoneList
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not app_exit Then
        Cancel = True
        Me.Hide
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
        
        picWeight_Paint
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
    DrawRect picWeight.hdc, 0, 0, w, h
    
    'weight
    Dim wt As Long
    wt = (w - 2) * skin_weight
    If wt > 0 Then
        picWeight.FillColor = RGB(200, 200, 200)
        picWeight.ForeColor = RGB(200, 200, 200)
        DrawRect picWeight.hdc, 1, 1, wt + 1, h - 1
    End If
    
    'slider
    picWeight.FillColor = RGB(200, 63, 63)
    picWeight.ForeColor = RGB(200, 63, 63)
    DrawRect picWeight.hdc, wt + 1 - 2, 0, wt + 1 + 2, h
    
End Sub

Private Sub FillBoneList()
    Me.lstBones.Clear
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
        Next i
    End With
End Sub
