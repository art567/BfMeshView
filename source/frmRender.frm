VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmRender 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Render Lighting"
   ClientHeight    =   6495
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6495
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRender.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   433
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   433
   Begin VB.Frame fraAntiAliasing 
      Caption         =   "Anti-Aliasing"
      Height          =   1095
      Left            =   4080
      TabIndex        =   49
      Top             =   1800
      Width           =   2295
      Begin VB.TextBox txtAAAThreshold 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         TabIndex        =   51
         Top             =   600
         Width           =   735
      End
      Begin VB.CheckBox chkAAAEnable 
         Caption         =   "Adaptive Anti-Aliasing"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   50
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label labMisc 
         AutoSize        =   -1  'True
         Caption         =   "Threshold:"
         Height          =   195
         Index           =   13
         Left            =   120
         TabIndex        =   52
         Top             =   600
         Width           =   765
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2520
      TabIndex        =   43
      Top             =   6000
      Width           =   1095
   End
   Begin VB.CommandButton cmdRender 
      Caption         =   "Render"
      Height          =   375
      Left            =   120
      TabIndex        =   42
      Top             =   6000
      Width           =   1095
   End
   Begin VB.CommandButton cmdPause 
      Caption         =   "Pause"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1320
      TabIndex        =   41
      Top             =   6000
      Width           =   1095
   End
   Begin VB.Frame fraFalloff 
      Caption         =   "Distance Falloff"
      Height          =   1455
      Left            =   4080
      TabIndex        =   29
      Top             =   3000
      Width           =   2295
      Begin VB.TextBox txtFalloffEnd 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         TabIndex        =   34
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox txtFalloffStart 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         TabIndex        =   32
         Top             =   600
         Width           =   735
      End
      Begin VB.CheckBox chkFalloffEnable 
         Caption         =   "Enabled"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   240
         Width           =   975
      End
      Begin VB.Label labMisc 
         AutoSize        =   -1  'True
         Caption         =   "Range End:"
         Height          =   195
         Index           =   10
         Left            =   120
         TabIndex        =   33
         Top             =   960
         Width           =   840
      End
      Begin VB.Label labMisc 
         AutoSize        =   -1  'True
         Caption         =   "Range Start:"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   31
         Top             =   600
         Width           =   930
      End
   End
   Begin VB.Frame fraAccelerate 
      Caption         =   "Acceleration"
      Height          =   1095
      Left            =   3000
      TabIndex        =   21
      Top             =   7080
      Width           =   2295
      Begin VB.TextBox txtAccelThreshold 
         Height          =   285
         Left            =   1440
         TabIndex        =   23
         Top             =   600
         Width           =   735
      End
      Begin VB.CheckBox chkAccelEnable 
         Caption         =   "Enabled"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   975
      End
      Begin VB.Label labMisc 
         AutoSize        =   -1  'True
         Caption         =   "Threshold:"
         Height          =   195
         Index           =   7
         Left            =   120
         TabIndex        =   24
         Top             =   600
         Width           =   765
      End
   End
   Begin VB.Frame fraProgress 
      Caption         =   "Progress"
      Height          =   1335
      Left            =   120
      TabIndex        =   18
      Top             =   4560
      Width           =   6255
      Begin ComctlLib.ProgressBar pgbTotal 
         Height          =   255
         Left            =   1200
         TabIndex        =   19
         Top             =   240
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   450
         _Version        =   327682
         Appearance      =   1
      End
      Begin ComctlLib.ProgressBar pgbPass 
         Height          =   255
         Left            =   1200
         TabIndex        =   37
         Top             =   600
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   450
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.Label labStats 
         Caption         =   "Samples: 0/0 Time taken: 0.0 sec"
         Height          =   195
         Left            =   1200
         TabIndex        =   40
         Top             =   960
         Width           =   4860
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Statistics:"
         Height          =   195
         Left            =   120
         TabIndex        =   39
         Top             =   960
         Width           =   705
      End
      Begin VB.Label labPass 
         AutoSize        =   -1  'True
         Caption         =   "Current Pass:"
         Height          =   195
         Left            =   120
         TabIndex        =   38
         Top             =   630
         Width           =   990
      End
      Begin VB.Label labTotal 
         AutoSize        =   -1  'True
         Caption         =   "Total:"
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   300
         Width           =   420
      End
   End
   Begin VB.Frame fraSettings 
      Caption         =   "Settings"
      Height          =   2655
      Left            =   120
      TabIndex        =   9
      Top             =   1800
      Width           =   3855
      Begin VB.ComboBox cbbResolution 
         Height          =   315
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   45
         Top             =   1440
         Width           =   1575
      End
      Begin VB.TextBox txtPasses 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         TabIndex        =   36
         Top             =   1800
         Width           =   735
      End
      Begin VB.TextBox txtFarPlane 
         Height          =   285
         Left            =   1200
         TabIndex        =   28
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox txtNearPlane 
         Height          =   285
         Left            =   1200
         TabIndex        =   27
         Top             =   1080
         Width           =   735
      End
      Begin VB.CheckBox chkTwoSided 
         Caption         =   "Two-Sided"
         Height          =   255
         Left            =   2160
         TabIndex        =   17
         Top             =   360
         Width           =   1095
      End
      Begin VB.CheckBox chkHemisphere 
         Caption         =   "Hemispherical"
         Height          =   255
         Left            =   2160
         TabIndex        =   16
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox txtPadding 
         Height          =   285
         Left            =   1200
         TabIndex        =   15
         Top             =   2160
         Width           =   735
      End
      Begin VB.TextBox txtFov 
         Height          =   285
         Left            =   1200
         TabIndex        =   14
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox txtFrameSize 
         Height          =   285
         Left            =   1200
         TabIndex        =   13
         Top             =   360
         Width           =   735
      End
      Begin VB.Label labMisc 
         AutoSize        =   -1  'True
         Caption         =   "Resolution:"
         Height          =   195
         Index           =   12
         Left            =   2160
         TabIndex        =   46
         Top             =   1110
         Width           =   810
      End
      Begin VB.Label labMisc 
         AutoSize        =   -1  'True
         Caption         =   "Passes:"
         Height          =   195
         Index           =   11
         Left            =   120
         TabIndex        =   35
         Top             =   1800
         Width           =   555
      End
      Begin VB.Label labMisc 
         AutoSize        =   -1  'True
         Caption         =   "Far Plane:"
         Height          =   195
         Index           =   9
         Left            =   120
         TabIndex        =   26
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label labMisc 
         AutoSize        =   -1  'True
         Caption         =   "Near Plane:"
         Height          =   195
         Index           =   8
         Left            =   120
         TabIndex        =   25
         Top             =   1080
         Width           =   840
      End
      Begin VB.Label labMisc 
         AutoSize        =   -1  'True
         Caption         =   "Padding:"
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   12
         Top             =   2160
         Width           =   630
      End
      Begin VB.Label labMisc 
         AutoSize        =   -1  'True
         Caption         =   "Field Of View:"
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   990
      End
      Begin VB.Label labMisc 
         AutoSize        =   -1  'True
         Caption         =   "Frame Size:"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   840
      End
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Browse..."
      Height          =   375
      Left            =   5160
      TabIndex        =   4
      Top             =   360
      Width           =   1095
   End
   Begin VB.Frame fraOutput 
      Caption         =   "Output"
      Height          =   1575
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   6255
      Begin VB.CheckBox chkOutputNormals 
         Caption         =   "Save Normal Map"
         Height          =   255
         Left            =   3960
         TabIndex        =   53
         Top             =   1080
         Width           =   1575
      End
      Begin VB.CheckBox chkShowOutput 
         Caption         =   "Show Output Progress"
         Height          =   255
         Left            =   3960
         TabIndex        =   48
         Top             =   720
         Width           =   1935
      End
      Begin VB.CheckBox chkOverwriteWarning 
         Caption         =   "Overwrite Warning"
         Height          =   255
         Left            =   1920
         TabIndex        =   47
         Top             =   720
         Width           =   2295
      End
      Begin VB.CheckBox chkOutputAlpha 
         Caption         =   "Save Alpha Channel"
         Enabled         =   0   'False
         Height          =   255
         Left            =   1920
         TabIndex        =   44
         Top             =   1080
         Width           =   1815
      End
      Begin VB.TextBox txtHeight 
         Height          =   285
         Left            =   960
         TabIndex        =   8
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox txtWidth 
         Height          =   285
         Left            =   960
         TabIndex        =   7
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox txtOutput 
         Height          =   285
         Left            =   960
         TabIndex        =   3
         Top             =   300
         Width           =   3975
      End
      Begin VB.Label labMisc 
         AutoSize        =   -1  'True
         Caption         =   "Height:"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   6
         Top             =   1110
         Width           =   525
      End
      Begin VB.Label labMisc 
         AutoSize        =   -1  'True
         Caption         =   "Width:"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   750
         Width           =   480
      End
      Begin VB.Label labMisc 
         AutoSize        =   -1  'True
         Caption         =   "Filename:"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   330
         Width           =   690
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   5280
      TabIndex        =   0
      Top             =   6000
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog cdlOutput 
      Left            =   4080
      Top             =   6000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmRender"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
        
    Me.cbbResolution.AddItem "1:1 (100%)": Me.cbbResolution.ItemData(0) = 1
    Me.cbbResolution.AddItem "1:2 (50%)": Me.cbbResolution.ItemData(1) = 2
    Me.cbbResolution.AddItem "1:4 (25%)": Me.cbbResolution.ItemData(2) = 4
    Me.cbbResolution.AddItem "1:8 (12.5%)": Me.cbbResolution.ItemData(3) = 8
    
    'set defaults
    Me.txtOutput.Text = lmoutput
    Me.chkOutputAlpha.Value = Abs(lmoutputalpha)
    Me.txtWidth.Text = lmwidth
    Me.txtHeight.Text = lmheight
    Me.cbbResolution.ListIndex = DivTwo(lmres)
    Me.chkOverwriteWarning.Value = Abs(lmwarnoverwrite)
    Me.chkShowOutput.Value = Abs(lmshowoutput)
    Me.chkOutputNormals.Value = Abs(lmoutputnormals)
    
    Me.txtFrameSize.Text = lmframesize
    Me.txtFov.Text = lmfov
    Me.txtNearPlane.Text = lmnear
    Me.txtFarPlane.Text = lmfar
    Me.txtPasses.Text = lmpasses
    Me.txtPadding.Text = lmpadding
    Me.chkTwoSided.Value = Abs(lmtwosided)
    Me.chkHemisphere.Value = Abs(lmhemisphere)
    
    Me.chkAAAEnable.Value = Abs(lmaaa): chkAAAEnable_Click
    Me.txtAAAThreshold.Text = lmaaathres
    
    Me.chkAccelEnable.Value = Abs(lmaccel): chkAccelEnable_Click
    Me.txtAccelThreshold.Text = lmaccelthres
    
    Me.chkFalloffEnable.Value = Abs(lmfalloff): chkFalloffEnable_Click
    Me.txtFalloffStart.Text = lmfalloffstart
    Me.txtFalloffEnd.Text = lmfalloffend
    
    Center Me
End Sub

Private Sub CopyGuiSettings()
    
    lmoutput = Me.txtOutput.Text
    lmoutputalpha = (Me.chkOutputAlpha.Value <> vbUnchecked)
    lmwidth = val(Me.txtWidth.Text)
    lmheight = val(Me.txtHeight.Text)
    lmwarnoverwrite = (Me.chkOverwriteWarning.Value <> vbUnchecked)
    lmshowoutput = (Me.chkShowOutput.Value <> vbUnchecked)
    lmoutputnormals = (Me.chkOutputNormals.Value <> vbUnchecked)
    
    lmres = Me.cbbResolution.ItemData(Me.cbbResolution.ListIndex)
    lmframesize = val(Me.txtFrameSize.Text)
    lmfov = val(Me.txtFov.Text)
    lmnear = val(Me.txtNearPlane.Text)
    lmfar = val(Me.txtFarPlane.Text)
    lmpasses = val(Me.txtPasses.Text)
    lmpadding = val(Me.txtPadding.Text)
    lmtwosided = (Me.chkTwoSided.Value <> vbUnchecked)
    lmhemisphere = (Me.chkHemisphere.Value <> vbUnchecked)
    
    lmaaa = (Me.chkAAAEnable.Value <> vbUnchecked)
    lmaaathres = val(Me.txtAAAThreshold.Text)
    
    lmaccel = (Me.chkAccelEnable.Value <> vbUnchecked)
    lmaccelthres = val(Me.txtAccelThreshold.Text)
    
    lmfalloff = (Me.chkFalloffEnable.Value <> vbUnchecked)
    lmfalloffstart = val(Me.txtFalloffStart.Text)
    lmfalloffend = val(Me.txtFalloffEnd.Text)
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload frmOutput
End Sub

Private Sub cmdBrowse_Click()
    With Me.cdlOutput
        .DialogTitle = "Output"
        .Filter = "TGA (*.tga)|*.tga"
        .FilterIndex = 1
        .flags = cdlOFNOverwritePrompt Or cdlOFNNoReadOnlyReturn
        .InitDir = current_folder
        .CancelError = True
        .filename = lmoutput
        On Error Resume Next
        .ShowSave
        If Not Err.Number = cdlCancel Then
            On Error GoTo 0
            
            Me.txtOutput.Text = .filename
        End If
    End With
End Sub

Private Sub chkAAAEnable_Click()
Dim b As Boolean
    b = (chkAAAEnable.Value <> vbUnchecked)
    EnableControl Me.txtAAAThreshold, b
End Sub

Private Sub chkAccelEnable_Click()
Dim b As Boolean
    b = (chkAccelEnable.Value <> vbUnchecked)
    EnableControl Me.txtAccelThreshold, b
End Sub

Private Sub chkFalloffEnable_Click()
Dim b As Boolean
    b = (Me.chkFalloffEnable.Value <> vbUnchecked)
    EnableControl Me.txtFalloffStart, b
    EnableControl Me.txtFalloffEnd, b
End Sub

Private Sub cmdRender_Click()
    
    If Not myobj.loaded Then
        MsgBox "No OBJ mesh loaded.", vbExclamation
        Exit Sub
    End If
    
    CopyGuiSettings
    SaveConfig app_configfile
    
    'overwrite warning
    If lmwarnoverwrite Then
        If FileExist(lmoutput) Then
            If Not MsgBox("Output file will be overwritten, continue?", vbYesNoCancel Or vbExclamation) = vbYes Then
                Exit Sub
            End If
        End If
        
    End If
    
    Me.cmdRender.Enabled = False
    Me.cmdPause.Enabled = True
    Me.cmdCancel.Enabled = True
    
    lmabort = False
    lmpause = False
        
    RenderLighting
    
    Me.pgbPass.Value = 0
    Me.pgbTotal.Value = 0
    Me.cmdRender.Enabled = True
    Me.cmdPause.Enabled = False
    Me.cmdCancel.Enabled = False
End Sub

Private Sub cmdPause_Click()
    lmpause = Not lmpause
    If lmpause Then
        Me.cmdPause.Caption = "Resume"
    Else
        Me.cmdPause.Caption = "Pause"
    End If
End Sub

Private Sub cmdCancel_Click()
    lmabort = True
    lmpause = False
End Sub

Private Sub cmdClose_Click()
    lmabort = True
    CopyGuiSettings
    SaveConfig app_configfile
    Unload Me
End Sub

' --- misc -------------------------------------------------------------

'...
Private Sub EnableControl(ByRef ctl As Control, ByVal b As Boolean)
    ctl.Enabled = b
    If b Then
        ctl.BackColor = &H80000005
    Else
        ctl.BackColor = &H8000000F
    End If
End Sub

