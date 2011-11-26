Attribute VB_Name = "a_Selection"
Option Explicit

Public toolmode As Long '0=view, 1=select
Private modelmatrix(15) As GLdouble
Private projectionmatrix(15) As GLdouble
Private viewport(3) As GLint


'updates matrices
Public Sub GetProjectionInfo()
    glGetDoublev GL_MODELVIEW_MATRIX, modelmatrix(0)
    glGetDoublev GL_PROJECTION_MATRIX, projectionmatrix(0)
    glGetIntegerv GL_VIEWPORT, viewport(0)
End Sub


'projects world space position to screen coordinate
Private Function Project(ByRef pos As float3) As float3
Dim winx As GLdouble
Dim winy As GLdouble
Dim winz As GLdouble
    gluProject pos.X, pos.Y, pos.z, _
               modelmatrix(0), projectionmatrix(0), viewport(0), _
               winx, winy, winz
    Project.X = winx
    Project.Y = viewport(3) - winy
    Project.z = winz
End Function


'selects vertex
Public Sub BF2SelectVert(ByVal minx As Single, ByVal miny As Single, ByVal maxx As Single, ByVal maxy As Single)
    With vmesh
        If Not .loadok Then Exit Sub
        'ClearVertSelection
        SetVertFlags2 selgeom, sellod
        
        Dim stride As Long
        stride = .vertstride / 4
        
        Dim i As Long
        For i = 0 To .vertnum - 1
            If .vertflag(i) Then
                
                'get vertex position
                Dim v As float3 'note: all DICE stuff is mirrored on X axis
                If vmesh.hasSkinVerts Then
                    v.X = -.skinvert(i).X
                    v.Y = .skinvert(i).Y
                    v.z = .skinvert(i).z
                Else
                    v.X = -.vert(i * stride + 0)
                    v.Y = .vert(i * stride + 1)
                    v.z = .vert(i * stride + 2)
                End If
                
                'project to screen
                Dim sv As float3
                sv = Project(v)
                
                'clear vert selection
                If Not frmMain.keyctrl And Not frmMain.keyalt Then
                    .vertsel(i) = 0
                End If
                If sv.z > 0 Then
                    If sv.X >= minx Then
                        If sv.X <= maxx Then
                            If sv.Y >= miny Then
                                If sv.Y <= maxy Then
                                    If frmMain.keyalt Then
                                        .vertsel(i) = 0
                                    Else
                                        .vertsel(i) = 1
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
                
            End If
        Next i
        
    End With
End Sub


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

