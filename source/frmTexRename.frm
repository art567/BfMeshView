VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTexRename 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Rename Textures"
   ClientHeight    =   4950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9375
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTexRename.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   330
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   625
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ImageList imlIcon 
      Left            =   8280
      Top             =   2280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   13
      ImageHeight     =   13
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTexRename.frx":1042
            Key             =   "tex"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTexRename.frx":1184
            Key             =   "texmissing"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Browse..."
      Height          =   375
      Left            =   8160
      TabIndex        =   4
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   375
      Left            =   8160
      TabIndex        =   3
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox txtRename 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   4560
      Width           =   7920
   End
   Begin MSComctlLib.ListView lstTextures 
      Height          =   4335
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   7646
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "imlIcon"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   952
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Name"
         Object.Width           =   13229
      EndProperty
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   8160
      TabIndex        =   0
      Top             =   600
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog cdlFile 
      Left            =   8400
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmTexRename"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mapnum As Long      'number of unique textures
Private mapname() As String 'array of texture paths
Private mapid() As Long     'array of texmap IDs


Private Sub Form_Load()
    lstTextures.ColumnHeaders.Item(2).width = Me.lstTextures.width - lstTextures.ColumnHeaders.Item(1).width - 6 - 16
    FillList
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdBrowse_Click()
    
    Dim fname As String
    Dim defpath As String
    Dim basepath As String
    
    'get current texture
    Dim i As Long
    For i = 1 To texmapnum
        If texmap(i).origrelfilename = Me.txtRename.Text Then
            fname = CleanFilePath(texmap(i).filename)
            Exit For
        End If
    Next i
    
    'determine base path
    If Not FileExist(fname) Then
        fname = ""
        defpath = vmesh.filename
    Else
    
        Dim oldname As String
        oldname = CleanFilePath(Me.txtRename.Text)
        
        If Len(oldname) > 0 Then
            
            'get base folder of old relative texture path
            Dim oldbase As String
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
            
            'subtract relative path
            Dim relpath As String
            relpath = Replace(newname, basepath, "")
            
            'copy path to textbox
            Me.txtRename.Text = relpath
        End If
    End With
    
End Sub

Private Sub txtRename_KeyDown(KeyCode As Integer, Shift As Integer)
    'RenameItem Me.lstTextures.SelectedItem
End Sub

Private Sub txtRename_Change()
    RenameItem Me.lstTextures.SelectedItem
End Sub

Private Sub RenameItem(ByRef itm As MSComctlLib.ListItem)
    If itm.SubItems(1) = Me.txtRename.Text Then Exit Sub
    itm.SubItems(1) = Me.txtRename.Text
    itm.ListSubItems(1).ForeColor = RGB(197, 113, 0)
End Sub

'Private Sub lstTextures_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'    lstTextures.SortKey = ColumnHeader.Index - 1
'    lstTextures.Sorted = True
'    lstTextures.SortOrder = 1 Xor lstTextures.SortOrder
'End Sub

Private Sub lstTextures_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Me.txtRename.Text = Item.SubItems(1)
End Sub

Private Sub FillList()
    
    'clear list
    Me.lstTextures.ListItems.Clear
    mapnum = 0
    Erase mapname()
    Erase mapid()
    
    'make list of unique textures
    With vmesh
        If Not .loadok Then Exit Sub
        Dim i As Long
        For i = 0 To .geomnum - 1
            With .geom(i)
                Dim j As Long
                For j = 0 To .lodnum - 1
                    With .lod(j)
                        Dim k As Long
                        For k = 0 To .matnum - 1
                            With .mat(k)
                                Dim m As Long
                                For m = 0 To .mapnum - 1
                                    
                                    Dim dupe As Boolean
                                    Dim q As Long
                                    dupe = False
                                    For q = 0 To mapnum - 1
                                        If mapname(q) = .map(m) Then
                                            dupe = True
                                        End If
                                    Next q
                                    If Not dupe Then
                                        mapnum = mapnum + 1
                                        ReDim Preserve mapname(0 To mapnum - 1)
                                        ReDim Preserve mapid(0 To mapnum - 1)
                                        mapname(mapnum - 1) = .map(m)
                                        mapid(mapnum - 1) = .texmapid(m)
                                    End If
                                    
                                Next m
                            End With
                        Next k
                    End With
                Next j
            End With
        Next i
    End With
    
    'fill list
    For i = 0 To mapnum - 1
        
        Dim itm As MSComctlLib.ListItem
        Set itm = Me.lstTextures.ListItems.Add(, "@" & i, i + 1)
        
        itm.SubItems(1) = mapname(i)
        
        'assign texture name color
        If mapid(i) = 0 Then
            itm.SmallIcon = 2
            itm.ListSubItems(1).ForeColor = RGB(127, 0, 0)
        Else
            itm.SmallIcon = 1
            itm.ListSubItems(1).ForeColor = RGB(0, 127, 0)
        End If
    Next i
    
End Sub

'applies texture rename
Private Sub cmdApply_Click()
    On Error GoTo errhandler
    
Dim i As Long
    
    'make list of new map names
    Dim newmapname() As String
    ReDim newmapname(0 To mapnum - 1)
    For i = 0 To mapnum - 1
        newmapname(i) = Me.lstTextures.ListItems(i + 1).SubItems(1)
    Next i
    
    'replace
    With vmesh
        If Not .loadok Then Exit Sub
        
        For i = 0 To .geomnum - 1
            With .geom(i)
                Dim j As Long
                For j = 0 To .lodnum - 1
                    With .lod(j)
                        Dim k As Long
                        For k = 0 To .matnum - 1
                            With .mat(k)
                                Dim m As Long
                                For m = 0 To .mapnum - 1
                                    
                                    Dim q As Long
                                    For q = 0 To mapnum - 1
                                        If .map(m) = mapname(q) Then
                                            
                                            .map(m) = newmapname(q)
                                            
                                        End If
                                    Next q
                                    
                                Next m
                            End With
                        Next k
                    End With
                Next j
            End With
        Next i
    End With
    
    'reload textures
    frmMain.mnuToolsReloadTextures_Click
    
    'close
    'Unload Me
    FillList
    
    'succes
    Exit Sub
errhandler:
    MsgBox "cmdApply_Click" & vbLf & err.Description, vbCritical
End Sub

