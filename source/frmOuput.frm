VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmOutput 
   Caption         =   "Render Output"
   ClientHeight    =   4125
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4890
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
   Icon            =   "frmOuput.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   275
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   326
   Begin ComctlLib.StatusBar stsMain 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   2
      Top             =   3840
      Width           =   4890
      _ExtentX        =   8625
      _ExtentY        =   503
      Style           =   1
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox picContainer 
      BackColor       =   &H8000000C&
      HasDC           =   0   'False
      Height          =   2175
      Left            =   120
      ScaleHeight     =   141
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   229
      TabIndex        =   0
      Top             =   120
      Width           =   3495
      Begin VB.PictureBox picMain 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   1455
         Left            =   480
         ScaleHeight     =   97
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   129
         TabIndex        =   1
         Top             =   240
         Visible         =   0   'False
         Width           =   1935
      End
   End
End
Attribute VB_Name = "frmOutput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function xSetPixel Lib "gdi32" Alias "SetPixel" (ByVal hdc As Long, _
                                                                 ByVal x As Long, _
                                                                 ByVal y As Long, _
                                                                 ByVal color As Long) As Long


Private sw As Long
Private sh As Long

Private Sub Form_Load()
     Center Me
End Sub

Private Sub Form_Resize()
Dim x As Long
Dim y As Long
    If Not Me.WindowState = vbMinimized Then
        
        If Me.width < 200 * 15 Then Me.width = 200 * 15
        If Me.height < 200 * 15 Then Me.height = 200 * 15
        
        Me.picContainer.Move 3, 3, Me.ScaleWidth - 6, Me.ScaleHeight - 6 - Me.stsMain.height
        
        x = (picContainer.ScaleWidth / 2) - (sw / 2)
        y = (picContainer.ScaleHeight / 2) - (sh / 2)
        If x < 0 Then x = 0
        If y < 0 Then y = 0
        
        Me.picMain.Move x, y
    End If
End Sub

Public Sub SetSize(ByVal w As Long, ByVal h As Long)
    
    If w <> sw Or h <> sh Then
        Me.picMain.Cls
        
        Dim ww As Long
        Dim wh As Long
        ww = w + 70
        wh = h + 70
        If ww < 200 Then ww = 200
        If wh < 200 Then wh = 200
        Me.width = ww * 15
        Me.height = wh * 15
        
    End If
    
    sw = w
    sh = h
    
    Me.picMain.width = sw
    Me.picMain.height = sh
    Me.picMain.Visible = True
    
    Me.stsMain.SimpleText = sw & "x" & sh
    
    Form_Resize
    DoEvents
End Sub

Public Sub SetPixel(ByVal x As Long, ByVal y As Long, ByRef c As Byte)
    'xSetPixel Me.picMain.hdc, x, y, RGB(c, c, c)
    Me.picMain.PSet (x, sh - y), RGB(c, c, c)
End Sub
