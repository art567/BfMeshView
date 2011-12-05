Attribute VB_Name = "a_Selection"
Option Explicit

Public toolmode As Long '0=view, 1=select
Private modelmatrix(15) As GLdouble
Private projmatrix(15) As GLdouble
Private viewport(3) As GLint


'updates matrices
Public Sub GetProjectionInfo()
    glGetDoublev GL_MODELVIEW_MATRIX, modelmatrix(0)
    glGetDoublev GL_PROJECTION_MATRIX, projmatrix(0)
    glGetIntegerv GL_VIEWPORT, viewport(0)
End Sub


'projects world space position to screen coordinate
Public Function Project(ByRef pos As float3) As float3
Dim winx As GLdouble
Dim winy As GLdouble
Dim winz As GLdouble
    gluProject pos.x, pos.y, pos.z, _
               modelmatrix(0), projmatrix(0), viewport(0), _
               winx, winy, winz
    Project.x = winx
    Project.y = viewport(3) - winy
    Project.z = winz
End Function


'unprojects screen space (pixels) position to world space
Public Function Unproject(ByRef pos As float3) As float3
Dim objx As GLdouble
Dim objy As GLdouble
Dim objz As GLdouble
    gluUnProject pos.x, viewport(3) - pos.y, pos.z, modelmatrix(0), projmatrix(0), viewport(0), objx, objy, objz
    Unproject.x = objx
    Unproject.y = objy
    Unproject.z = objz
End Function


'unprojects screen space (float) position to world space
Public Function Unproject2(ByRef pos As float3) As float3
Dim objx As GLdouble
Dim objy As GLdouble
Dim objz As GLdouble
Dim vp(3) As GLint
    vp(0) = 0
    vp(1) = 0
    vp(2) = 1
    vp(3) = 1
    gluUnProject pos.x, vp(3) - pos.y, pos.z, modelmatrix(0), projmatrix(0), vp(0), objx, objy, objz
    Unproject2.x = objx
    Unproject2.y = objy
    Unproject2.z = objz
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
                    v.x = -.skinvert(i).x
                    v.y = .skinvert(i).y
                    v.z = .skinvert(i).z
                Else
                    v.x = -.vert(i * stride + 0)
                    v.y = .vert(i * stride + 1)
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
                    If sv.x >= minx Then
                        If sv.x <= maxx Then
                            If sv.y >= miny Then
                                If sv.y <= maxy Then
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

