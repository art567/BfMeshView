Attribute VB_Name = "a_Config"
Option Explicit

Public app_configfile As String
Public app_exit As Boolean

Public current_folder As String
Public opt_runmaximized As Boolean
Public opt_loadtextures As Boolean
Public opt_loadsamples As Boolean
Public opt_loadcon As Boolean
Public opt_loadviewsettings As Boolean
Public opt_uselocaltexpath As Boolean
Public opt_useglsl As Boolean
Public opt_loadspeclut As Boolean
Public bgcolor As color4f

Private Type texpath_type
    use As Boolean
    path As String
End Type

Public texpathnum As Long
Public texpath() As texpath_type

Public suffixnum As Long
Public suffix() As String
Public suffix_sel As Long

Public nmap_lastinput As String
Public nmap_lastoutput As String
Public nmap_padding As Long
Public nmap_flatten As Single


'loads default configuration
Public Sub LoadDefaultConfig()
    
    view_wire = False
    view_lighting = True
    view_textures = True
    view_poly = True
    view_verts = False
    view_edges = False
    view_normals = False
    view_tangents = False
    view_backfaces = False
    view_bounds = False
    view_bonesys = False
    view_samples = True
    view_axis = False
    view_grids = True
    view_envmap = True
    view_bumpmap = True
    
    bgcolor.r = 0.25
    bgcolor.g = 0.25
    bgcolor.b = 0.25
    bgcolor.a = 0
    
    opt_runmaximized = False
    opt_loadtextures = True
    opt_loadsamples = False
    opt_loadcon = False
    opt_loadviewsettings = False
    opt_useglsl = True
    opt_loadspeclut = False
    
    current_folder = App.path
    
End Sub


'loads configuration from file
Public Sub LoadConfig(ByRef filename As String)
Dim ff As Integer
Dim ln As String
Dim str() As String
Dim linenum As Long
Dim b As Long
Dim c As Long
Dim i As Long
Dim j As Long
Dim skip As Boolean
    
    On Error GoTo errorhandler
    
    'check if file exists
    If Not FileExist(filename) Then
        'MsgBox "File " & Chr(34) & filename & Chr(34) & " not found.", vbExclamation
        Exit Sub
    End If
    
    'open file
    ff = FreeFile
    Open filename For Input As #ff
    Do Until EOF(ff)
        linenum = linenum + 1
        Line Input #ff, ln
        
        'remove whitespaces
        ln = Trim$(ln)
        
        'skip comments
        skip = False
        If Len(ln) = 0 Then skip = True
        If Left$(ln, 1) = ";" Then skip = True
        If Left$(ln, 1) = "[" Then skip = True
        
        'process
        If Not skip Then
            str() = Split(ln, "=", 2)
            Select Case str(0)
            
            'misc
            Case "lastpath":         current_folder = str(1)
            Case "loadtextures":     opt_loadtextures = val(str(1))
            Case "loadsamples":      opt_loadsamples = val(str(1))
            Case "loadcon":          opt_loadcon = val(str(1))
            Case "loadviewsettings": opt_loadviewsettings = val(str(1))
            Case "uselocaltexpath":  opt_uselocaltexpath = val(str(1))
            Case "useglsl":          opt_useglsl = val(str(1))
            Case "loadspeclut":      opt_loadspeclut = val(str(1))
            Case "runmaximized":     opt_runmaximized = val(str(1))
            Case "bgcolor":
                str() = Split(str(1), "/")
                bgcolor.r = val(str(0)) / 255
                bgcolor.g = val(str(1)) / 255
                bgcolor.b = val(str(2)) / 255
                
            'texpath
            Case "texpath"
                str() = Split(str(1), ",", 2)
                
                texpathnum = texpathnum + 1
                ReDim Preserve texpath(1 To texpathnum)
                texpath(texpathnum).use = val(str(0))
                texpath(texpathnum).path = str(1)
                
            'suffix
            Case "suffix"
                suffixnum = suffixnum + 1
                ReDim Preserve suffix(1 To suffixnum)
                suffix(suffixnum) = str(1)
                
            'lightmap renderer
            Case "lmoutput":        lmoutput = str(1)
            Case "lmoutputalpha":   lmoutputalpha = val(str(1))
            Case "lmwidth":         lmwidth = val(str(1))
            Case "lmheight":        lmheight = val(str(1))
            Case "lmwarnoverwrite": lmwarnoverwrite = val(str(1))
            Case "lmshowoutput":    lmshowoutput = val(str(1))
            
            Case "lmres":           lmres = val(str(1))
            Case "lmframesize":     lmframesize = val(str(1))
            Case "lmfov":           lmfov = val(str(1))
            Case "lmnear":          lmnear = val(str(1))
            Case "lmfar":           lmfar = val(str(1))
            Case "lmpasses":        lmpasses = val(str(1))
            Case "lmpadding":       lmpadding = val(str(1))
            Case "lmtwosided":      lmtwosided = val(str(1))
            Case "lmhemisphere":    lmhemisphere = val(str(1))
            
            Case "lmaaa":           lmaaa = val(str(1))
            Case "lmaaathres":      lmaaathres = val(str(1))
            
            Case "lmaccel":         lmaccel = val(str(1))
            Case "lmaccelthres":    lmaccelthres = val(str(1))
            
            Case "lmfalloff":       lmfalloff = val(str(1))
            Case "lmfalloffstart":  lmfalloffstart = val(str(1))
            Case "lmfalloffend":    lmfalloffend = val(str(1))
            
            'normal map converter
            Case "nmap_lastinput":  nmap_lastinput = str(1)
            Case "nmap_lastoutput": nmap_lastoutput = str(1)
            Case "nmap_padding":    nmap_padding = val(str(1))
            Case "nmap_flatten": nmap_flatten = val(str(1))
            
            End Select
            
            'view settings
            If opt_loadviewsettings Then
                Select Case str(0)
                Case "view_wire":       view_wire = val(str(1))
                Case "view_lighting":   view_lighting = val(str(1))
                Case "view_textures":   view_textures = val(str(1))
                Case "view_poly":       view_poly = val(str(1))
                Case "view_verts":      view_verts = val(str(1))
                Case "view_edges":      view_edges = val(str(1))
                Case "view_normals":    view_normals = val(str(1))
                Case "view_backfaces":  view_backfaces = val(str(1))
                Case "view_tangents":   view_tangents = val(str(1))
                Case "view_bounds":     view_bounds = val(str(1))
                Case "view_bonesys":    view_bonesys = val(str(1))
                Case "view_samples":    view_samples = val(str(1))
                Case "view_axis":       view_axis = val(str(1))
                Case "view_grids":      view_grids = val(str(1))
                Case "view_envmap":     view_envmap = val(str(1))
                Case "view_bumpmap":    view_bumpmap = val(str(1))

                End Select
            End If
            
        End If
        
    Loop
    
    'close file
    Close ff
    
    'success
    On Error GoTo 0
    Exit Sub
    
    'error
errorhandler:
    On Error GoTo 0
    MsgBox "LoadConfig" & vbCrLf & err.description, vbCritical
    MsgBox "Error on line " & linenum & ":" & vbLf & ln, vbInformation
    Close ff
End Sub


'saves configuration to file
Public Sub SaveConfig(ByRef filename As String)
Dim ff As Integer
Dim i As Long
Dim backup As Boolean
    
    On Error GoTo errorhandler
    
    'back up old file
    If FileExist(filename) Then
        FileCopy filename, filename & ".bak"
        Kill filename
        backup = True
    End If
    
    'open file
    ff = FreeFile
    Open filename For Output As #ff
    
    'print header
    Print #ff, "; " & App.Title & " Configuration File"
    Print #ff, ""
    
    'misc
    Print #ff, "[Misc]"
    Print #ff, "lastpath=" & GetFilePath(current_folder)
    Print #ff, "loadtextures=" & Abs(opt_loadtextures)
    Print #ff, "loadsamples=" & Abs(opt_loadsamples)
    Print #ff, "loadcon=" & Abs(opt_loadcon)
    Print #ff, "loadviewsettings=" & Abs(opt_loadviewsettings)
    Print #ff, "uselocaltexpath=" & Abs(opt_uselocaltexpath)
    Print #ff, "useglsl=" & Abs(opt_useglsl)
    Print #ff, "loadspeclut=" & Abs(opt_loadspeclut)
    Print #ff, "runmaximized=" & Abs(opt_runmaximized)
    Print #ff, "bgcolor=" & Round(bgcolor.r * 255) & "/" & Round(bgcolor.g * 255) & "/" & Round(bgcolor.b * 255)
    Print #ff, ""
    
    'print paths
    Print #ff, "[Texture Paths]"
    For i = 1 To texpathnum
        Print #ff, "texpath=" & Abs(texpath(i).use) & "," & texpath(i).path
    Next i
    For i = 1 To suffixnum
        Print #ff, "suffix=" & suffix(i)
    Next i
    Print #ff, ""
    
    'view settings
    Print #ff, "[View]"
    Print #ff, "view_wire=" & Abs(view_wire)
    Print #ff, "view_lighting=" & Abs(view_lighting)
    Print #ff, "view_textures=" & Abs(view_textures)
    Print #ff, "view_poly=" & Abs(view_poly)
    Print #ff, "view_verts=" & Abs(view_verts)
    Print #ff, "view_edges=" & Abs(view_edges)
    Print #ff, "view_normals=" & Abs(view_normals)
    Print #ff, "view_backfaces=" & Abs(view_backfaces)
    Print #ff, "view_tangents=" & Abs(view_tangents)
    Print #ff, "view_bounds=" & Abs(view_bounds)
    Print #ff, "view_bonesys=" & Abs(view_bonesys)
    Print #ff, "view_samples=" & Abs(view_samples)
    Print #ff, "view_axis=" & Abs(view_axis)
    Print #ff, "view_grids=" & Abs(view_grids)
    Print #ff, "view_envmap=" & Abs(view_envmap)
    Print #ff, "view_bumpmap=" & Abs(view_bumpmap)
    Print #ff, ""
    
    'lighting renderer
    Print #ff, "[Lighting Renderer]"
    Print #ff, "lmoutput=" & lmoutput
    Print #ff, "lmoutputalpha=" & Abs(lmoutputalpha)
    Print #ff, "lmwidth=" & lmwidth
    Print #ff, "lmheight=" & lmheight
    Print #ff, "lmwarnoverwrite=" & Abs(lmwarnoverwrite)
    Print #ff, "lmshowoutput=" & Abs(lmshowoutput)
    
    Print #ff, "lmres=" & lmres
    Print #ff, "lmframesize=" & lmframesize
    Print #ff, "lmfov=" & lmfov
    Print #ff, "lmnear=" & lmnear
    Print #ff, "lmfar=" & lmfar
    Print #ff, "lmpasses=" & lmpasses
    Print #ff, "lmpadding=" & lmpadding
    Print #ff, "lmtwosided=" & Abs(lmtwosided)
    Print #ff, "lmhemisphere=" & Abs(lmhemisphere)
    
    Print #ff, "lmaaa=" & Abs(lmaaa)
    Print #ff, "lmaaathres=" & val(lmaaathres)
    
    Print #ff, "lmaccel=" & Abs(lmaccel)
    Print #ff, "lmaccelthres=" & val(lmaccelthres)
    
    Print #ff, "lmfalloff=" & Abs(lmfalloff)
    Print #ff, "lmfalloffstart=" & lmfalloffstart
    Print #ff, "lmfalloffend=" & lmfalloffend
    Print #ff, ""
    
    'normal map converter
    Print #ff, "[Normal Map Converter]"
    Print #ff, "nmap_lastinput=" & nmap_lastinput
    Print #ff, "nmap_lastoutput=" & nmap_lastoutput
    Print #ff, "nmap_padding=" & nmap_padding
    Print #ff, "nmap_flatten=" & nmap_flatten
    Print #ff, ""
    
    'print footer
    Print #ff, "; end of file"
    
    'close file
    Close ff
    
    If backup Then
        Kill filename & ".bak"
    End If
    
    On Error GoTo 0
    Exit Sub
errorhandler:
    On Error GoTo 0
    MsgBox "SaveConfig" & vbCrLf & err.description, vbCritical
    Close ff
End Sub

