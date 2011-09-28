Attribute VB_Name = "BF2_Anim"
Option Explicit

Private Type baf_frame 'per-frame transform
    rot As quat
    pos As float3
End Type

Private Type baf_bonedata 'per-bone track
    datasize As Integer
    frame() As baf_frame
End Type

Public Type baf_file
    version As Long
    bonenum As Integer
    boneId() As Integer
    framenum As Long
    precision As Byte
    boneData() As baf_bonedata
    
    'internal
    filename As String
    loaded As Boolean
End Type

Public bf2baf As baf_file


'loads bf2 animation file
Public Function LoadBF2Anim(ByVal filename As String) As Boolean
Dim i As Long
    
    'unload
    UnloadBF2Anim
    
    If Not FileExist(filename) Then
        MsgBox "File " & filename & " not found.", vbExclamation
        Exit Function
    End If
    
    'open file
    Dim ff As Integer
    ff = FreeFile
    Open filename For Binary Access Read Lock Write As #ff
    
    With bf2baf
        .filename = filename
        .loaded = False
        
        'version (4 bytes)
        Get #ff, , .version
        Echo "version: " & .version
        
        'bonenum (2 bytes)
        Get #ff, , .bonenum
        Echo "bonenum: " & .bonenum
        
        'boneid (2 bytes * bonenum)
        ReDim .boneId(0 To .bonenum - 1)
        Get #ff, , .boneId()
        For i = 0 To .bonenum - 1
            Echo "boneID[" & i & "]: " & .boneId(i)
        Next i
        
        'framenum (4 bytes)
        Get #ff, , .framenum
        Echo "framenum: " & .framenum
        
        'precision (1 byte)
        Get #ff, , .precision
        Echo "precision: " & .precision
        
        'per-bone data
        Echo "bone data block start at " & loc(ff)
        ReDim .boneData(0 To .bonenum - 1)
        For i = 0 To .bonenum - 1
            Echo " bone " & i & " data start at " & loc(ff)
            
            If Not ReadBoneData(ff, .boneData(i)) Then
                Echo "ERROR OCCURRED, ABORTING"
                Exit For
            End If
            
            Echo " bone " & i & " data end at " & loc(ff)
        Next i
        Echo "bone data block end at " & loc(ff)
        
        .loaded = True
    End With
    
    Echo "done at: " & loc(ff)
    Echo "filesize: " & LOF(ff)
    
    'close file
    Close ff
    
    'show animation toolbar
    frmMain.picAnim.Visible = True
    
    'success
    LoadBF2Anim = True
    Exit Function
errhandler:
    MsgBox "LoadBF2Anim" & vbLf & err.description, vbCritical
End Function


'unloads file
Public Sub UnloadBF2Anim()
    With bf2baf
        .loaded = False
        .filename = ""
        
        'hide animation toolbar
        frmMain.picAnim.Visible = False
        
    End With
End Sub


'reads bone data
Private Function ReadBoneData(ByRef ff As Integer, ByRef boneData As baf_bonedata) As Boolean
    Dim j As Long
    Dim m As Long
    
    With boneData
        
        'datasize (2 bytes)
        Get #ff, , .datasize
        Echo "  datasize: " & .datasize
        
        'allocate frames
        ReDim .frame(0 To bf2baf.framenum - 1)
        
        'read streams
        '* 7 streams of 16-bit floating point values
        '* each stream corresponds to a transform axis
        '* first four are XYZW that store quaternion rotation
        '* next three are XYZ that store the position
        '* each stream may be RLE compressed
        '* each 16-bit float needs to be unpacked
        '* rotation values have fixed compression
        '* position values have variable compression
        
        For j = 1 To 7
            
            'keep track of current output frame
            Dim curFrame As Long
            curFrame = 0
            
            Dim dataLeft As Integer
            Get #ff, , dataLeft
            
            While (dataLeft > 0)
                
                'head (1 byte)
                Dim head As Byte
                Get #ff, , head
                
                'bit 8 is RLE compression flag
                Dim rleCompression As Boolean
                rleCompression = GetByteBit(head, 8)
                
                'determine number of frames
                'stored in the leftmost 7 bits, so we have to zero our bit 8
                Dim numframes As Long
                SetByteBit head, 8, False
                numframes = head
                
                'offset to next header (2 bytes)
                Dim nextHeader As Byte
                Get #ff, , nextHeader
                
                'value (2 bytes)
                Dim value As Integer
                
                If rleCompression Then Get #ff, , value 'outside loop
                For m = 0 To numframes - 1
                    If Not rleCompression Then Get #ff, , value 'inside loop
                    
                    If j = 1 Then .frame(curFrame).rot.X = DecompFloat(value, 15)
                    If j = 2 Then .frame(curFrame).rot.Y = DecompFloat(value, 15)
                    If j = 3 Then .frame(curFrame).rot.z = DecompFloat(value, 15)
                    If j = 4 Then .frame(curFrame).rot.w = DecompFloat(value, 15)
                    
                    If j = 5 Then .frame(curFrame).pos.X = DecompFloat(value, bf2baf.precision)
                    If j = 6 Then .frame(curFrame).pos.Y = DecompFloat(value, bf2baf.precision)
                    If j = 7 Then .frame(curFrame).pos.z = DecompFloat(value, bf2baf.precision)
                    
                    curFrame = curFrame + 1
                Next m
                
                'decrement
                dataLeft = dataLeft - nextHeader
                
            Wend
            
        Next j
        
        'fix quaternion rotations
        For j = 0 To bf2baf.framenum - 1
            With .frame(j)
                .rot.X = -.rot.X
                .rot.Y = -.rot.Y
                .rot.z = -.rot.z
            End With
        Next j
        
    End With
    
    'success
    ReadBoneData = True
End Function


'converts variable precision 16-bit float to float
Private Function DecompFloat(ByRef tmpInt As Integer, ByRef precision As Byte) As Single
    
    'seems to convert from unsigned to signed? I think we don't have to do this in VB
    'Dim tmpVal As Integer
    'tmpVal = tmpInt
    'If tmpVal > 32767 Then
    '    tmpVal = tmpVal - 65535
    'End If
    
    Dim flt16_mult As Single
    flt16_mult = CSng(32767) / (2 ^ (15 - precision)) 'maybe swap??
    
    DecompFloat = CSng(tmpInt) / flt16_mult
End Function


'--- bit utility functions ---------------------------------------------------------------


'gets bit at position
Private Function GetByteBit(ByRef b As Byte, ByRef pos As Byte) As Boolean
    ASSERT pos >= 1 And pos <= 8, "GetByteBit > pos out of range"
    GetByteBit = ((b And (2 ^ (pos - 1))) > 0)
End Function


'sets bit at position
Private Sub SetByteBit(ByRef b As Byte, ByRef pos As Byte, ByRef val As Boolean)
    ASSERT pos >= 1 And pos <= 8, "GetByteBit > pos out of range"
    If val Then
        b = b Or (2 ^ (pos - 1))
    Else
        b = b And Not (2 ^ (pos - 1))
    End If
End Sub


'shifts bits to right
Function BitShiftRight(ByVal v As Long, ByVal Shift As Integer) As Long
    BitShiftRight = v \ 2 ^ Shift
End Function


'shifts bits to left
Function BitShiftLeft(ByVal v As Long, ByVal Shift As Integer) As Long
    BitShiftLeft = v * 2 ^ Shift
End Function


'--- treeview ---------------------------------------------------------------------------------------------

'fill treeview
Public Sub FillTreeBF2Anim(ByRef tree As MSComctlLib.TreeView)
    With bf2baf
        On Error GoTo errhandler
        If Not .loaded Then Exit Sub
        
        Dim n As MSComctlLib.node
        
        'file root
        Dim rootname As String
        rootname = "baf_root"
        Set n = tree.Nodes.Add(, , rootname, GetFileName(.filename), "file")
        n.Expanded = True
        n.tag = 0
        
        'version leaf
        Set n = tree.Nodes.Add(rootname, tvwChild, rootname & "|ver", "Version: " & .version, "prop")
        n.tag = 0
        
        'nodenum leaf
        Set n = tree.Nodes.Add(rootname, tvwChild, rootname & "|framenum", "Frames: " & .framenum, "prop")
        n.tag = 0
        
        'nodenum leaf
        Set n = tree.Nodes.Add(rootname, tvwChild, rootname & "|nodenum", "Nodes: " & .bonenum, "prop")
        n.tag = 0
        
        'nodeIDs leaf
        Set n = tree.Nodes.Add(rootname, tvwChild, "nodeIDs", "NodeIDs", "geom")
        n.tag = 0
        
        'loop nodes
        Dim i As Long
        For i = 0 To .bonenum - 1
            
            Set n = tree.Nodes.Add("nodeIDs", tvwChild, "BAFnodeID" & i, "Node " & .boneId(i), "lod")
            n.Expanded = True
            n.tag = i
            
        Next i
        
    End With
    Exit Sub
errhandler:
    MsgBox "FillTreeBF2Anim" & vbLf & err.description, vbCritical
End Sub

