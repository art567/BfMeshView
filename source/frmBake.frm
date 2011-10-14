VERSION 5.00
Begin VB.Form frmBake 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Bake Texture"
   ClientHeight    =   2535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4335
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBake.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   169
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   289
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtUvChannel 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   "5"
      Top             =   1755
      Width           =   855
   End
   Begin VB.ComboBox cbbGeom 
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
      ItemData        =   "frmBake.frx":000C
      Left            =   1200
      List            =   "frmBake.frx":000E
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
   Begin VB.ComboBox cbbLod 
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
      ItemData        =   "frmBake.frx":0010
      Left            =   1200
      List            =   "frmBake.frx":0012
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   540
      Width           =   1815
   End
   Begin VB.TextBox txtHeight 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2160
      TabIndex        =   4
      Top             =   1380
      Width           =   855
   End
   Begin VB.ComboBox cbbPreset 
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
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   960
      Width           =   1815
   End
   Begin VB.TextBox txtWidth 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1200
      TabIndex        =   3
      Top             =   1380
      Width           =   855
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
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
      Left            =   3120
      TabIndex        =   8
      Top             =   600
      Width           =   1095
   End
   Begin VB.TextBox txtPadding 
      Height          =   285
      Left            =   1200
      TabIndex        =   6
      Text            =   "6"
      Top             =   2130
      Width           =   855
   End
   Begin VB.CommandButton cmdBake 
      Caption         =   "Bake"
      Height          =   375
      Left            =   3120
      TabIndex        =   7
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label labMisc 
      AutoSize        =   -1  'True
      Caption         =   "UV Channel:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   120
      TabIndex        =   14
      Top             =   1815
      Width           =   885
   End
   Begin VB.Label labMisc 
      AutoSize        =   -1  'True
      Caption         =   "Geom:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   13
      Top             =   180
      Width           =   465
   End
   Begin VB.Label labMisc 
      AutoSize        =   -1  'True
      Caption         =   "LOD Level:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   4
      Left            =   120
      TabIndex        =   12
      Top             =   600
      Width           =   780
   End
   Begin VB.Label labMisc 
      AutoSize        =   -1  'True
      Caption         =   "Size Preset:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   5
      Left            =   120
      TabIndex        =   11
      Top             =   1020
      Width           =   855
   End
   Begin VB.Label labMisc 
      AutoSize        =   -1  'True
      Caption         =   "Custom Size:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   10
      Top             =   1425
      Width           =   930
   End
   Begin VB.Label labMisc 
      AutoSize        =   -1  'True
      Caption         =   "Padding:"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   9
      Top             =   2190
      Width           =   630
   End
End
Attribute VB_Name = "frmBake"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    
    Dim i As Long
    With vmesh
        For i = 0 To .geomnum - 1
            Me.cbbGeom.AddItem "Geom " & i
        Next i
        Me.cbbGeom.ListIndex = 0
    End With
    
    Me.cbbPreset.AddItem "4096x4096"
    Me.cbbPreset.AddItem "2048x2048"
    Me.cbbPreset.AddItem "1024x1024"
    Me.cbbPreset.AddItem "512x512"
    Me.cbbPreset.AddItem "256x256"
    Me.cbbPreset.AddItem "128x128"
    Me.cbbPreset.ListIndex = 2
    
    Me.txtUvChannel.Text = vmesh.uvnum
    
End Sub

Private Sub cmdBake_Click()
    
    Me.MousePointer = vbHourglass
    Me.Enabled = False
    Me.cmdBake.Enabled = False
    
    Dim geom As Long
    geom = Me.cbbGeom.ListIndex
    
    Dim lod As Long
    lod = Me.cbbLod.ListIndex
    
    Dim w As Long
    Dim h As Long
    w = val(Me.txtWidth.Text)
    h = val(Me.txtHeight.Text)
    
    Dim uvchan As Long
    uvchan = val(Me.txtUvChannel.Text) - 1
    If uvchan < 0 Then uvchan = 0
    If uvchan > vmesh.uvnum - 1 Then uvchan = vmesh.uvnum - 1
    Me.txtUvChannel.Text = uvchan + 1
    
    Dim padding As Long
    padding = val(Me.txtPadding.Text)
    If padding < 0 Then padding = 0
    If padding > 64 Then padding = 0
    Me.txtPadding.Text = padding
    
    BakeTexture geom, lod, w, h, uvchan, padding
    
    Me.cmdBake.Enabled = True
    Me.Enabled = True
    Me.MousePointer = vbDefault
    
End Sub

Private Sub cmdClose_Click()
    Unload Me
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

Private Sub cbbPreset_Click()
    Dim name As String
    name = Me.cbbPreset.List(cbbPreset.ListIndex)
    
    Dim str() As String
    str() = Split(name, "x")
    
    Me.txtWidth.Text = str(0)
    Me.txtHeight.Text = str(1)
End Sub

