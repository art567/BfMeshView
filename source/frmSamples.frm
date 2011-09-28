VERSION 5.00
Begin VB.Form frmSamples 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Generate Samples"
   ClientHeight    =   2655
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
   Icon            =   "frmSamples.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   177
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   289
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkIgnoreErrors 
      Caption         =   "Ignore Errors"
      Height          =   195
      Left            =   1200
      TabIndex        =   14
      Top             =   2340
      Width           =   1335
   End
   Begin VB.CheckBox chkEdgeMargin 
      Caption         =   "Clamp To Triangle Edge"
      Height          =   195
      Left            =   1200
      TabIndex        =   13
      Top             =   2040
      Value           =   1  'Checked
      Width           =   2175
   End
   Begin VB.TextBox txtPadding 
      Height          =   285
      Left            =   1200
      TabIndex        =   11
      Text            =   "6"
      Top             =   1680
      Width           =   855
   End
   Begin VB.TextBox txtUvChan 
      Height          =   285
      Left            =   1200
      TabIndex        =   10
      Text            =   "5"
      Top             =   1320
      Width           =   855
   End
   Begin VB.ComboBox cbbLod 
      Height          =   315
      ItemData        =   "frmSamples.frx":000C
      Left            =   1200
      List            =   "frmSamples.frx":000E
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   120
      Width           =   1815
   End
   Begin VB.ComboBox cbbPreset 
      Height          =   315
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   540
      Width           =   1815
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Close"
      Height          =   375
      Left            =   3120
      TabIndex        =   3
      Top             =   600
      Width           =   1095
   End
   Begin VB.TextBox txtHeight 
      Height          =   285
      Left            =   2160
      TabIndex        =   1
      Top             =   960
      Width           =   855
   End
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "Generate"
      Height          =   375
      Left            =   3120
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox txtWidth 
      Height          =   285
      Left            =   1200
      TabIndex        =   0
      Top             =   960
      Width           =   855
   End
   Begin VB.Label labMisc 
      AutoSize        =   -1  'True
      Caption         =   "Padding:"
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   12
      Top             =   1725
      Width           =   630
   End
   Begin VB.Label labMisc 
      AutoSize        =   -1  'True
      Caption         =   "UV Channel:"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   9
      Top             =   1365
      Width           =   885
   End
   Begin VB.Label labMisc 
      AutoSize        =   -1  'True
      Caption         =   "Size Preset:"
      Height          =   195
      Index           =   5
      Left            =   120
      TabIndex        =   8
      Top             =   615
      Width           =   855
   End
   Begin VB.Label labMisc 
      AutoSize        =   -1  'True
      Caption         =   "LOD Level:"
      Height          =   195
      Index           =   4
      Left            =   120
      TabIndex        =   7
      Top             =   180
      Width           =   780
   End
   Begin VB.Label labMisc 
      AutoSize        =   -1  'True
      Caption         =   "Custom Size:"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   1005
      Width           =   930
   End
End
Attribute VB_Name = "frmSamples"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    
    Me.cbbPreset.AddItem "1024x1024"
    Me.cbbPreset.AddItem "512x512"
    Me.cbbPreset.AddItem "256x256"
    Me.cbbPreset.AddItem "128x128"
    Me.cbbPreset.AddItem "64x64"
    Me.cbbPreset.AddItem "32x32"
    Me.cbbPreset.AddItem "16x16"
    Me.cbbPreset.ListIndex = 2
    
    If vmesh.loadok Then
        If vmesh.geomnum > 0 Then
            Dim i As Long
            For i = 0 To vmesh.geom(0).lodnum - 1
                Me.cbbLod.AddItem "LOD " & i
            Next i
            Me.cbbLod.ListIndex = 0
            
            Me.cmdGenerate.Enabled = True
            Me.cbbLod.Enabled = True
            Me.cbbLod.BackColor = &H80000005
        End If
    End If
End Sub

Private Sub cbbPreset_Click()
    Dim name As String
    name = Me.cbbPreset.List(cbbPreset.ListIndex)
    
    Dim str() As String
    str() = Split(name, "x")
    
    Me.txtWidth.Text = str(0)
    Me.txtHeight.Text = str(1)
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdGenerate_Click()
    
    Dim L As Long
    Dim w As Long
    Dim h As Long
    Dim c As Long
    Dim p As Long
    L = Me.cbbLod.ListIndex
    w = val(Me.txtWidth.Text)
    h = val(Me.txtHeight.Text)
    c = val(Me.txtUvChan.Text) - 1
    p = val(Me.txtPadding.Text)
    
    Dim e As Boolean
    e = Me.chkEdgeMargin.value = vbChecked
    SAMP_IgnoreTriErrors = Me.chkIgnoreErrors.value = vbChecked
    
    WriteSamplesFile L, w, h, c, p, e
    
End Sub


Private Sub txtHeight_GotFocus()
    SelectOnFocus Me.txtHeight
End Sub

Private Sub txtWidth_GotFocus()
    SelectOnFocus Me.txtWidth
End Sub

Private Sub txtUvChan_GotFocus()
    SelectOnFocus Me.txtUvChan
End Sub

Private Sub txtPadding_GotFocus()
    SelectOnFocus Me.txtPadding
End Sub

Private Sub SelectOnFocus(ByRef txt As TextBox)
    txt.SelStart = 0
    txt.SelLength = Len(txt.Text)
End Sub
