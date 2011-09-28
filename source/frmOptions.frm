VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Preferences"
   ClientHeight    =   6375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6615
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   425
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   441
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog cdlColor 
      Left            =   720
      Top             =   5880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame fraMisc 
      Caption         =   "Miscellaneous"
      Height          =   735
      Left            =   120
      TabIndex        =   25
      Top             =   5040
      Width           =   6375
      Begin VB.CommandButton cmdReset 
         Caption         =   "Reset"
         Height          =   375
         Left            =   5160
         TabIndex        =   31
         Top             =   240
         Width           =   1095
      End
      Begin VB.PictureBox picBgColor 
         BackColor       =   &H00404040&
         HasDC           =   0   'False
         Height          =   255
         Left            =   2280
         ScaleHeight     =   13
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   13
         TabIndex        =   29
         Top             =   300
         Width           =   255
      End
      Begin VB.CheckBox chkMaximized 
         Caption         =   "Maximize On Start"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   300
         Width           =   1695
      End
      Begin VB.Label labMisc 
         AutoSize        =   -1  'True
         Caption         =   "Viewport Background Color"
         Height          =   195
         Left            =   2640
         TabIndex        =   30
         Top             =   300
         Width           =   1935
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "File Associations"
      Height          =   2055
      Left            =   120
      TabIndex        =   9
      Top             =   2880
      Width           =   6375
      Begin VB.CheckBox chkFileAssoc 
         Caption         =   ".bfmv"
         Height          =   255
         Index           =   17
         Left            =   120
         TabIndex        =   32
         Top             =   240
         Width           =   1095
      End
      Begin VB.CheckBox chkFileAssoc 
         Caption         =   ".baf"
         Height          =   255
         Index           =   16
         Left            =   1800
         TabIndex        =   28
         Top             =   240
         Width           =   1335
      End
      Begin VB.CheckBox chkFileAssoc 
         Caption         =   ".occ"
         Height          =   255
         Index           =   15
         Left            =   3480
         TabIndex        =   27
         Top             =   600
         Width           =   615
      End
      Begin VB.CheckBox chkFileAssoc 
         Caption         =   ".samp_03"
         Height          =   255
         Index           =   7
         Left            =   1800
         TabIndex        =   24
         Top             =   1680
         Width           =   1095
      End
      Begin VB.CheckBox chkFileAssoc 
         Caption         =   ".samp_02"
         Height          =   255
         Index           =   6
         Left            =   1800
         TabIndex        =   23
         Top             =   1320
         Width           =   1095
      End
      Begin VB.CheckBox chkFileAssoc 
         Caption         =   ".samp_01"
         Height          =   255
         Index           =   5
         Left            =   1800
         TabIndex        =   22
         Top             =   960
         Width           =   1095
      End
      Begin VB.CheckBox chkFileAssoc 
         Caption         =   ".samples"
         Height          =   255
         Index           =   0
         Left            =   1800
         TabIndex        =   21
         Top             =   600
         Width           =   975
      End
      Begin VB.CheckBox chkFileAssoc 
         Caption         =   ".obj"
         Height          =   255
         Index           =   11
         Left            =   4920
         TabIndex        =   20
         Top             =   240
         Width           =   735
      End
      Begin VB.CheckBox chkFileAssoc 
         Caption         =   ".rig"
         Height          =   255
         Index           =   8
         Left            =   4920
         TabIndex        =   19
         Top             =   1320
         Width           =   735
      End
      Begin VB.CheckBox chkFileAssoc 
         Caption         =   ".ske"
         Height          =   255
         Index           =   14
         Left            =   3480
         TabIndex        =   18
         Top             =   240
         Width           =   615
      End
      Begin VB.CheckBox chkFileAssoc 
         Caption         =   ".collisionmesh"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   17
         Top             =   1680
         Width           =   1335
      End
      Begin VB.CheckBox chkFileAssoc 
         Caption         =   ".tri"
         Height          =   255
         Index           =   9
         Left            =   4920
         TabIndex        =   16
         Top             =   960
         Width           =   735
      End
      Begin VB.CheckBox chkFileAssoc 
         Caption         =   ".geo"
         Height          =   255
         Index           =   10
         Left            =   4920
         TabIndex        =   15
         Top             =   600
         Width           =   735
      End
      Begin VB.CheckBox chkFileAssoc 
         Caption         =   ".tm"
         Height          =   255
         Index           =   12
         Left            =   3480
         TabIndex        =   14
         Top             =   1320
         Width           =   615
      End
      Begin VB.CheckBox chkFileAssoc 
         Caption         =   ".sm"
         Height          =   255
         Index           =   13
         Left            =   3480
         TabIndex        =   13
         Top             =   960
         Width           =   615
      End
      Begin VB.CheckBox chkFileAssoc 
         Caption         =   ".skinnedmesh"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   12
         Top             =   1320
         Width           =   1335
      End
      Begin VB.CheckBox chkFileAssoc 
         Caption         =   ".bundledmesh"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   11
         Top             =   960
         Width           =   1335
      End
      Begin VB.CheckBox chkFileAssoc 
         Caption         =   ".staticmesh"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5400
      TabIndex        =   5
      Top             =   5880
      Width           =   1095
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   375
      Left            =   4200
      TabIndex        =   4
      Top             =   5880
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Texture Paths"
      Height          =   2715
      Left            =   90
      TabIndex        =   0
      Top             =   60
      Width           =   6435
      Begin VB.CommandButton cmdEdit 
         Caption         =   "Edit..."
         Height          =   375
         Left            =   5190
         TabIndex        =   8
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CommandButton cmdPathDown 
         Caption         =   "Move Down"
         Height          =   375
         Left            =   5190
         TabIndex        =   7
         Top             =   2160
         Width           =   1095
      End
      Begin VB.CommandButton cmdPathUp 
         Caption         =   "Move Up"
         Height          =   375
         Left            =   5190
         TabIndex        =   6
         Top             =   1680
         Width           =   1095
      End
      Begin VB.CommandButton cmdPathRemove 
         Caption         =   "Remove"
         Height          =   375
         Left            =   5190
         TabIndex        =   3
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton cmdPathAdd 
         Caption         =   "Add..."
         Height          =   375
         Left            =   5190
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
      Begin VB.ListBox lstPaths 
         Height          =   2310
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   1
         Top             =   240
         Width           =   4935
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type BrowseInfo
    hWndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As String
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type

Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetActiveWindow Lib "user32" () As Long


Private Function SelectFolder(ByVal default As String) As String
    Const BIF_RETURNONLYFSDIRS = 1
    
    Dim lpIDList As Long
    Dim lResult As Long
    Dim sPath As String
    Dim tBrowse As BrowseInfo
    
    With tBrowse
        .hWndOwner = Me.hWnd
        .lpszTitle = "Select Folder"
        .ulFlags = BIF_RETURNONLYFSDIRS 'return only if the user selected a directory
    End With
    
    lpIDList = SHBrowseForFolder(tBrowse)
    If lpIDList Then
        sPath = String$(260, 0)
        
        SHGetPathFromIDList lpIDList, sPath
        CoTaskMemFree lpIDList
        
        SelectFolder = safeStr(sPath)
    Else
        SelectFolder = default
    End If
End Function

Private Sub Form_Load()
    FillPathList
    
    Dim i As Long
    With Me.chkFileAssoc
        For i = .LBound To .UBound
            
            If GetFileAssoc(.Item(i).Caption) Then
                .Item(i).value = vbChecked
            Else
                .Item(i).value = vbUnchecked
            End If
        Next i
    End With
    
    If opt_runmaximized Then
        Me.chkMaximized.value = vbChecked
    Else
        Me.chkMaximized.value = vbUnchecked
    End If
    
    Me.picBgColor.BackColor = RGB(bgcolor.r * 255, bgcolor.g * 255, bgcolor.b * 255)
    
End Sub

Private Sub cmdApply_Click()
    
    'save config
    Dim i As Long
    For i = 1 To texpathnum
        texpath(i).use = Me.lstPaths.Selected(i - 1)
        texpath(i).path = Me.lstPaths.List(i - 1)
    Next i
    SaveConfig app_configfile
    
    'apply file associations
    With Me.chkFileAssoc
        For i = .LBound To .UBound
            SetFileAssoc .Item(i).Caption, (.Item(i).value = vbChecked)
        Next i
    End With
    
    'run maximized
    opt_runmaximized = Me.chkMaximized.value
    
    'viewport background color
    bgcolor.r = CSng(Me.picBgColor.BackColor And &HFF&) / 255
    bgcolor.g = CSng((Me.picBgColor.BackColor And &HFF00&) \ 256) / 255
    bgcolor.b = CSng((Me.picBgColor.BackColor And &HFF0000) \ 65536) / 255
    
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPathAdd_Click()
Dim s As String
    's = InputBox("Path:", , "")
    s = SelectFolder("")
    If Len(s) > 0 Then
        texpathnum = texpathnum + 1
        ReDim Preserve texpath(1 To texpathnum)
        texpath(texpathnum).use = True
        texpath(texpathnum).path = s
    End If
    FillPathList
End Sub

Private Sub cmdPathRemove_Click()
Dim sel As Long
    sel = Me.lstPaths.ListIndex + 1
    If sel < 1 Then Exit Sub
    If sel > texpathnum Then Exit Sub
    
    'swap toward end of list
    If sel < texpathnum Then
        Dim i As Long
        For i = sel + 1 To texpathnum
            SwapPath i, i - 1
        Next i
    End If
    
    'delete
    texpathnum = texpathnum - 1
    If texpathnum > 0 Then
        ReDim Preserve texpath(1 To texpathnum)
    Else
        Erase texpath()
    End If
    
    FillPathList
    Me.lstPaths.ListIndex = -1
End Sub

Private Sub cmdEdit_Click()
Dim sel As Long
    sel = Me.lstPaths.ListIndex + 1
    If sel < 1 Then Exit Sub
    If sel > texpathnum Then Exit Sub
    
Dim s As String
    's = SelectFolder(texpath(sel).path) '
    s = InputBox("Edit Path:", "Edit Path", texpath(sel).path)
    If Len(s) > 0 Then
        texpath(sel).path = s
    End If
    FillPathList
End Sub

Private Sub cmdPathUp_Click()
Dim sel As Long
    sel = Me.lstPaths.ListIndex + 1
    If sel <= 1 Then Exit Sub
    If sel > texpathnum Then Exit Sub
    SwapPath sel - 1, sel
    FillPathList
    Me.lstPaths.ListIndex = sel - 1 - 1
End Sub

Private Sub cmdPathDown_Click()
Dim sel As Long
    sel = Me.lstPaths.ListIndex + 1
    If sel < 1 Then Exit Sub
    If sel >= texpathnum Then Exit Sub
    SwapPath sel, sel + 1
    FillPathList
    Me.lstPaths.ListIndex = sel + 1 - 1
End Sub

Private Sub FillPathList()
Dim i As Long
    Me.lstPaths.Clear
    For i = 1 To texpathnum
        Me.lstPaths.AddItem texpath(i).path
        Me.lstPaths.Selected(Me.lstPaths.NewIndex) = texpath(i).use
    Next i
    Me.lstPaths.ListIndex = -1
End Sub

Private Sub SwapPath(ByVal a As Long, ByVal b As Long)
Dim t_use As Boolean
Dim t_path As String
    t_use = texpath(a).use
    t_path = texpath(a).path
    texpath(a).use = texpath(b).use
    texpath(a).path = texpath(b).path
    texpath(b).use = t_use
    texpath(b).path = t_path
End Sub


Private Sub picBgColor_Click()
    PickColor picBgColor
End Sub

Public Sub PickColor(ByRef pic As PictureBox)
    With cdlColor
        .DialogTitle = "Pick Color"
        .color = pic.BackColor
        .flags = cdlCCRGBInit
        .CancelError = True
        On Error Resume Next
        .ShowColor
        If err.Number <> cdlCancel Then
            On Error GoTo 0
            pic.BackColor = .color
        End If
        On Error GoTo 0
    End With
End Sub

Private Sub cmdReset_Click()
    Me.picBgColor.BackColor = RGB(63, 63, 63)
End Sub

