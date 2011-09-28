VERSION 5.00
Begin VB.Form frmGLInfo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "OpenGL Info"
   ClientHeight    =   4050
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5415
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmGLInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   270
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   361
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdFindNext 
      Caption         =   "Find Next"
      Height          =   375
      Left            =   4365
      TabIndex        =   8
      Top             =   3600
      Width           =   975
   End
   Begin VB.TextBox txtSearch 
      Height          =   285
      Left            =   660
      TabIndex        =   6
      Top             =   3645
      Width           =   3555
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   4365
      TabIndex        =   5
      Top             =   60
      Width           =   975
   End
   Begin VB.ListBox lstExtensions 
      Height          =   2400
      Left            =   60
      TabIndex        =   4
      Top             =   1140
      Width           =   5295
   End
   Begin VB.Label labSearch 
      AutoSize        =   -1  'True
      Caption         =   "Search:"
      Height          =   195
      Left            =   60
      TabIndex        =   7
      Top             =   3690
      Width           =   555
   End
   Begin VB.Label labMisc 
      AutoSize        =   -1  'True
      Caption         =   "Supported extensions:"
      Height          =   195
      Left            =   60
      TabIndex        =   3
      Top             =   870
      Width           =   1635
   End
   Begin VB.Label labVersion 
      Caption         =   "Version:"
      Height          =   195
      Left            =   60
      TabIndex        =   2
      Top             =   600
      Width           =   5235
   End
   Begin VB.Label labRenderer 
      Caption         =   "Renderer:"
      Height          =   195
      Left            =   60
      TabIndex        =   1
      Top             =   330
      Width           =   4185
   End
   Begin VB.Label labVendor 
      Caption         =   "Vendor:"
      Height          =   195
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   4155
   End
   Begin VB.Menu mnuHidden 
      Caption         =   "Hidden"
      Visible         =   0   'False
      Begin VB.Menu mnuHiddenCopy 
         Caption         =   "Copy Selected"
      End
      Begin VB.Menu mnuHiddenCopyAll 
         Caption         =   "Copy All"
         Begin VB.Menu mnuHiddenCopyAllSpace 
            Caption         =   "Space Seperated"
         End
         Begin VB.Menu mnuHiddenCopyAllComma 
            Caption         =   "Comma Seperated"
         End
         Begin VB.Menu mnuHiddenCopyAllLine 
            Caption         =   "Line Seperated"
         End
      End
   End
End
Attribute VB_Name = "frmGLInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private extstr As String

Private Sub Form_Load()
Dim ptr As Long
Dim str As String
Dim arr() As String
Dim i As Integer
    ptr = glGetString(GL_VENDOR)
    str = CharToString(ptr)
    Me.labVendor.Caption = "Vendor: " & str
    
    ptr = glGetString(GL_RENDERER)
    str = CharToString(ptr)
    Me.labRenderer.Caption = "Renderer: " & str
    
    ptr = glGetString(GL_VERSION)
    str = CharToString(ptr)
    Me.labVersion.Caption = "Version: " & str
    
    ptr = glGetString(GL_EXTENSIONS)
    str = CharToString(ptr)
    
    extstr = Trim(str)
    
    Me.lstExtensions.Clear
    arr() = Split(str, " ")
    For i = LBound(arr()) To UBound(arr()) - 1
        Me.lstExtensions.AddItem arr(i)
    Next i
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub lstExtensions_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
        PopupMenu mnuHidden, , (x / 15) + Me.lstExtensions.Left, (y / 15) + Me.lstExtensions.Top
    End If
End Sub

Private Sub mnuHiddenCopy_Click()
    Clipboard.Clear
    Clipboard.SetText Me.lstExtensions.Text
End Sub

Private Sub mnuHiddenCopyAllSpace_Click()
    Clipboard.Clear
    Clipboard.SetText extstr
End Sub

Private Sub mnuHiddenCopyAllComma_Click()
    Clipboard.Clear
    Clipboard.SetText Replace(extstr, " ", ", ")
End Sub

Private Sub mnuHiddenCopyAllLine_Click()
    Clipboard.Clear
    Clipboard.SetText Replace(extstr, " ", vbCrLf)
End Sub

Private Sub txtSearch_Change()
    FindExt Me.txtSearch, 1
End Sub

Private Sub cmdFindNext_Click()
    FindExt Me.txtSearch, Me.lstExtensions.ListIndex + 1
End Sub

Private Sub FindExt(ByRef str As String, ByRef start As Integer)
Dim i As Long
    If Len(str) = 0 Then Exit Sub
    If i > Me.lstExtensions.ListCount Then Exit Sub
    For i = start To Me.lstExtensions.ListCount
        If InStr(1, LCase(Me.lstExtensions.List(i)), LCase(str)) Then
            Me.lstExtensions.ListIndex = i
            Exit Sub
        End If
    Next i
    Me.lstExtensions.ListIndex = -1
End Sub
