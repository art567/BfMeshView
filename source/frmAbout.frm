VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3495
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HasDC           =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   201
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   233
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1200
      TabIndex        =   0
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label labMisc 
      AutoSize        =   -1  'True
      Caption         =   "BAF parser based on code by Rexman"
      Height          =   195
      Index           =   1
      Left            =   375
      TabIndex        =   5
      Top             =   2160
      Width           =   2745
   End
   Begin VB.Label labMisc 
      Alignment       =   2  'Center
      Caption         =   "For Forgotten Hope"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   3240
   End
   Begin VB.Label labFHmod 
      Alignment       =   2  'Center
      Caption         =   "http://www.fhmod.org"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   120
      MouseIcon       =   "frmAbout.frx":000C
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   1770
      Width           =   3225
   End
   Begin VB.Label labLink 
      Alignment       =   2  'Center
      Caption         =   "placeholder"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   120
      MouseIcon       =   "frmAbout.frx":015E
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   1080
      Width           =   3225
   End
   Begin VB.Label labTitle 
      Caption         =   "placeholder"
      Height          =   855
      Left            =   720
      TabIndex        =   1
      Top             =   120
      Width           =   2655
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   120
      Picture         =   "frmAbout.frx":02B0
      Stretch         =   -1  'True
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_Load()
    Me.labLink.Caption = App.Comments
    Me.labTitle = App.Title & vbCrLf & _
                  "Version " & App.Major & "." & App.Minor & "." & App.Revision & vbCrLf & vbCrLf & _
                  App.LegalCopyright
End Sub

Private Sub cmdOk_Click()
    Unload Me
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.labLink.ForeColor = &HFF0000
    Me.labFHmod.ForeColor = &HFF0000
End Sub


Private Sub labLink_Click()
    On Error GoTo errorhandler
    
    ShellExecute Me.hWnd, "open", Me.labLink.Caption, vbNullString, "", 0
    
    On Error GoTo 0
    Exit Sub
errorhandler:
End Sub
Private Sub labLink_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.labLink.ForeColor = &HFF&
End Sub


Private Sub labFHmod_Click()
    On Error GoTo errorhandler
    
    ShellExecute Me.hWnd, "open", Me.labFHmod.Caption, vbNullString, "", 0
    
    On Error GoTo 0
    Exit Sub
errorhandler:
End Sub
Private Sub labFHmod_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.labFHmod.ForeColor = &HFF&
End Sub


