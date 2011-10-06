Attribute VB_Name = "a_Math"
Option Explicit

Public Const PI As Single = 3.14159265358979
Public Const RADTODEG As Single = (180 / PI)
Public Const DEGTORAD As Single = (PI / 180)

Public Type float2
    x As Single
    y As Single
End Type

Public Type float3
    x As Single
    y As Single
    z As Single
End Type

Public Type float4
    x As Single
    y As Single
    z As Single
    w As Single
End Type

Public Type quat
    x As Single
    y As Single
    z As Single
    w As Single
End Type

Public Type matrix4
    m(0 To 15) As Single
End Type

Public Sub mat4identity(ByRef m As matrix4)
    m.m(0) = 1
    m.m(1) = 0
    m.m(2) = 0
    m.m(3) = 0
    
    m.m(4) = 0
    m.m(5) = 1
    m.m(6) = 0
    m.m(7) = 0
    
    m.m(8) = 0
    m.m(9) = 0
    m.m(10) = 1
    m.m(11) = 0
    
    m.m(12) = 0
    m.m(13) = 0
    m.m(14) = 0
    m.m(15) = 1
End Sub

Public Sub mat4setrot(ByRef m As matrix4, ByRef q As quat)
    m.m(0) = (1 - (2 * ((q.y * q.y) + (q.z * q.z))))
    m.m(1) = (2 * ((q.x * q.y) + (q.z * q.w)))
    m.m(2) = (2 * ((q.x * q.z) - (q.y * q.w)))
    m.m(3) = 0
    m.m(4) = (2 * ((q.x * q.y) - (q.z * q.w)))
    m.m(5) = (1 - (2 * ((q.x * q.x) + (q.z * q.z))))
    m.m(6) = (2 * ((q.y * q.z) + (q.x * q.w)))
    m.m(7) = 0
    m.m(8) = (2 * ((q.x * q.z) + (q.y * q.w)))
    m.m(9) = (2 * ((q.y * q.z) - (q.x * q.w)))
    m.m(10) = (1 - (2 * ((q.x * q.x) + (q.y * q.y))))
    m.m(11) = 0
End Sub

Public Sub mat4setrotYXZ(ByRef m As matrix4, ByRef r As float3)
    Dim cx As Single
    Dim sx As Single
    Dim cy As Single
    Dim sy As Single
    Dim cz As Single
    Dim sz As Single
    
    cx = Cos(r.x * DEGTORAD)
    sx = Sin(r.x * DEGTORAD)
    cy = Cos(r.y * DEGTORAD)
    sy = Sin(r.y * DEGTORAD)
    cz = Cos(r.z * DEGTORAD)
    sz = Sin(r.z * DEGTORAD)
    
    m.m(0) = sx * sy * sz + cy * cz
    m.m(4) = sx * sy * cz - cy * sz
    m.m(8) = cx * sy
    
    m.m(1) = cx * sz
    m.m(5) = cx * cz
    m.m(9) = -sx
    
    m.m(2) = -sy * cz + sx * cy * sz
    m.m(6) = sy * sz + sx * cy * cz
    m.m(10) = cx * cy
End Sub

Public Sub mat4setpos(ByRef m As matrix4, ByRef pos As float3)
    m.m(12) = pos.x
    m.m(13) = pos.y
    m.m(14) = pos.z
End Sub

Public Function mat4getpos(ByRef m As matrix4) As float3
    mat4getpos.x = m.m(12)
    mat4getpos.y = m.m(13)
    mat4getpos.z = m.m(14)
End Function

Public Function mat4rotvec(ByRef m As matrix4, ByRef v As float3) As float3
    mat4rotvec.x = (m.m(0) * v.x + m.m(4) * v.y + m.m(8) * v.z)
    mat4rotvec.y = (m.m(1) * v.x + m.m(5) * v.y + m.m(9) * v.z)
    mat4rotvec.z = (m.m(2) * v.x + m.m(6) * v.y + m.m(10) * v.z)
End Function

Public Function mat4transvec(ByRef m As matrix4, ByRef v As float3) As float3
    mat4transvec.x = (m.m(0) * v.x + m.m(4) * v.y + m.m(8) * v.z) + m.m(12)
    mat4transvec.y = (m.m(1) * v.x + m.m(5) * v.y + m.m(9) * v.z) + m.m(13)
    mat4transvec.z = (m.m(2) * v.x + m.m(6) * v.y + m.m(10) * v.z) + m.m(14)
End Function

Public Function mat4mult(ByRef a As matrix4, ByRef b As matrix4) As matrix4
    mat4mult.m(0) = a.m(0) * b.m(0) + a.m(1) * b.m(4) + a.m(2) * b.m(8) + a.m(3) * b.m(12)
    mat4mult.m(1) = a.m(0) * b.m(1) + a.m(1) * b.m(5) + a.m(2) * b.m(9) + a.m(3) * b.m(13)
    mat4mult.m(2) = a.m(0) * b.m(2) + a.m(1) * b.m(6) + a.m(2) * b.m(10) + a.m(3) * b.m(14)
    mat4mult.m(3) = a.m(0) * b.m(3) + a.m(1) * b.m(7) + a.m(2) * b.m(11) + a.m(3) * b.m(15)
    
    mat4mult.m(4) = a.m(4) * b.m(0) + a.m(5) * b.m(4) + a.m(6) * b.m(8) + a.m(7) * b.m(12)
    mat4mult.m(5) = a.m(4) * b.m(1) + a.m(5) * b.m(5) + a.m(6) * b.m(9) + a.m(7) * b.m(13)
    mat4mult.m(6) = a.m(4) * b.m(2) + a.m(5) * b.m(6) + a.m(6) * b.m(10) + a.m(7) * b.m(14)
    mat4mult.m(7) = a.m(4) * b.m(3) + a.m(5) * b.m(7) + a.m(6) * b.m(11) + a.m(7) * b.m(15)
    
    mat4mult.m(8) = a.m(8) * b.m(0) + a.m(9) * b.m(4) + a.m(10) * b.m(8) + a.m(11) * b.m(12)
    mat4mult.m(9) = a.m(8) * b.m(1) + a.m(9) * b.m(5) + a.m(10) * b.m(9) + a.m(11) * b.m(13)
    mat4mult.m(10) = a.m(8) * b.m(2) + a.m(9) * b.m(6) + a.m(10) * b.m(10) + a.m(11) * b.m(14)
    mat4mult.m(11) = a.m(8) * b.m(3) + a.m(9) * b.m(7) + a.m(10) * b.m(11) + a.m(11) * b.m(15)
    
    mat4mult.m(12) = a.m(12) * b.m(0) + a.m(13) * b.m(4) + a.m(14) * b.m(8) + a.m(15) * b.m(12)
    mat4mult.m(13) = a.m(12) * b.m(1) + a.m(13) * b.m(5) + a.m(14) * b.m(9) + a.m(15) * b.m(13)
    mat4mult.m(14) = a.m(12) * b.m(2) + a.m(13) * b.m(6) + a.m(14) * b.m(10) + a.m(15) * b.m(14)
    mat4mult.m(15) = a.m(12) * b.m(3) + a.m(13) * b.m(7) + a.m(14) * b.m(11) + a.m(15) * b.m(15)
End Function

Public Sub mat4lookat(ByRef m As matrix4, ByRef dir As float3, ByRef up As float3)
    Dim vx As float3
    Dim vy As float3
    Dim vz As float3
    
    vz.x = -dir.x
    vz.y = -dir.y
    vz.z = -dir.z
    vz = Normalize(vz)
    vx = Normalize(CrossProduct(up, vz))    'x = up cross z
    vy = CrossProduct(vz, vx)               ' y = z cross x
    
    m.m(0) = vx.x
    m.m(1) = vx.y
    m.m(2) = vx.z
    m.m(4) = vy.x
    m.m(5) = vy.y
    m.m(6) = vy.z
    m.m(8) = vz.x
    m.m(9) = vz.y
    m.m(10) = vz.z
End Sub

Public Function float2(ByVal x As Single, ByVal y As Single) As float2
    float2.x = x
    float2.y = y
End Function

Public Function float3(ByVal x As Single, ByVal y As Single, ByVal z As Single) As float3
    float3.x = x
    float3.y = y
    float3.z = z
End Function

Public Function float4(ByVal x As Single, ByVal y As Single, ByVal z As Single, ByVal w As Single) As float4
    float4.x = x
    float4.y = y
    float4.z = z
    float4.w = w
End Function

Public Function DotProduct(ByRef vector1 As float3, ByRef vector2 As float3) As Single
    DotProduct = (vector1.x * vector2.x + vector1.y * vector2.y + vector1.z * vector2.z)
End Function

Public Function CrossProduct(ByRef vec1 As float3, ByRef vec2 As float3) As float3
    CrossProduct.x = (vec1.y * vec2.z) - (vec1.z * vec2.y)
    CrossProduct.y = (vec1.z * vec2.x) - (vec1.x * vec2.z)
    CrossProduct.z = (vec1.x * vec2.y) - (vec1.y * vec2.x)
End Function

Public Function Distance(ByRef pos1 As float3, ByRef pos2 As float3) As Single
    Distance = Sqr((pos1.x - pos2.x) ^ 2 + (pos1.y - pos2.y) ^ 2 + (pos1.z - pos2.z) ^ 2)
End Function

Public Function IsPowerOfTwo(ByVal value As Long) As Boolean
    If (value < 2) Then Exit Function
    If (value And (value - 1)) = False Then IsPowerOfTwo = True
End Function

Public Function Floor(ByVal value As Single)
    Floor = Fix(value)
End Function

Public Function Ceil(ByVal value As Single)
    Ceil = Round(value + 0.5)
End Function

Public Function min(ByVal a As Single, ByVal b As Single) As Single
    If a < b Then
        min = a
    Else
        min = b
    End If
End Function

Public Function max(ByVal a As Single, ByVal b As Single) As Single
    If a > b Then
        max = a
    Else
        max = b
    End If
End Function

'add
Public Function AddFloat3(ByRef a As float3, ByRef b As float3) As float3
    AddFloat3.x = a.x + b.x
    AddFloat3.y = a.y + b.y
    AddFloat3.z = a.z + b.z
End Function

'subtract
Public Function SubFloat3(ByRef a As float3, ByRef b As float3) As float3
    SubFloat3.x = a.x - b.x
    SubFloat3.y = a.y - b.y
    SubFloat3.z = a.z - b.z
End Function

'multiply
Public Function ScaleFloat3(ByRef v As float3, ByVal s As Single) As float3
    ScaleFloat3.x = v.x * s
    ScaleFloat3.y = v.y * s
    ScaleFloat3.z = v.z * s
End Function

Public Function Clamp(ByVal v As Single, ByVal vmin As Single, ByVal vmax As Single) As Single
    If v < vmin Then v = vmin
    If v > vmax Then v = vmax
    Clamp = v
End Function

Public Function Rotate(ByRef Center As float3, ByRef position As float3, ByRef rotation As float3) As float3
Dim tmp As float3
Dim pos As float3
Dim rot As float3
    rot.x = rotation.x / (180 / PI)
    rot.y = rotation.y / (180 / PI)
    rot.z = rotation.z / (180 / PI)
    pos.x = position.x - Center.x
    pos.y = position.y - Center.y
    pos.z = position.z - Center.z
    tmp.y = (pos.y * Cos(rot.x)) - (pos.z * Sin(rot.x))
    tmp.z = (pos.z * Cos(rot.x)) + (pos.y * Sin(rot.x))
    pos.y = tmp.y
    pos.z = tmp.z
    tmp.z = (pos.z * Cos(rot.y)) - (pos.x * Sin(rot.y))
    tmp.x = (pos.x * Cos(rot.y)) + (pos.z * Sin(rot.y))
    pos.z = tmp.z
    pos.x = tmp.x
    tmp.x = (pos.x * Cos(rot.z)) - (pos.y * Sin(rot.z))
    tmp.y = (pos.y * Cos(rot.z)) + (pos.x * Sin(rot.z))
    Rotate.x = tmp.x + Center.x
    Rotate.y = tmp.y + Center.y
    Rotate.z = tmp.z + Center.z
End Function

Public Sub identityMatrix(ByRef matrix() As Single)
    matrix(0) = 1: matrix(1) = 0: matrix(2) = 0
    matrix(3) = 0: matrix(4) = 1: matrix(5) = 0
    matrix(6) = 0: matrix(7) = 0: matrix(8) = 1
End Sub

Public Sub rotateMatrix(ByRef matrix() As Single, ByRef rot As float3)
Static cr As Single
Static sr As Single
Static cp As Single
Static sp As Single
Static cy As Single
Static sy As Single
Static srsp As Single
Static crsp As Single
    
    cr = Cos(PI * rot.x / 180)
    sr = Sin(PI * rot.x / 180)
    cp = Cos(PI * rot.y / 180)
    sp = Sin(PI * rot.y / 180)
    cy = Cos(PI * rot.z / 180)
    sy = Sin(PI * rot.z / 180)

    matrix(0) = cp * cy
    matrix(1) = cp * sy
    matrix(2) = -sp
    
    srsp = sr * sp
    crsp = cr * sp
    
    matrix(3) = srsp * cy - cr * sy
    matrix(4) = srsp * sy + cr * cy
    matrix(5) = sr * cp
    
    matrix(6) = crsp * cy + sr * sy
    matrix(7) = crsp * sy - sr * cy
    matrix(8) = cr * cp
End Sub

Public Sub transposeMat4x4(ByRef m() As Single)
Dim i As Long
Dim j As Long
Dim t(0 To 3, 0 To 3) As Single
    
    'copy old
    For i = 0 To 3
        For j = 0 To 3
            t(i, j) = m(i, j)
        Next j
    Next i
    
    'transpose
    m(0, 0) = t(0, 0)
    m(0, 1) = t(1, 0)
    m(0, 2) = t(2, 0)
    m(0, 3) = t(3, 0)
    
    m(1, 0) = t(0, 1)
    m(1, 1) = t(1, 1)
    m(1, 2) = t(2, 1)
    m(1, 3) = t(3, 1)
    
    m(2, 0) = t(0, 2)
    m(2, 1) = t(1, 2)
    m(2, 2) = t(2, 2)
    m(2, 3) = t(3, 2)
    
    m(3, 0) = t(0, 3)
    m(3, 1) = t(1, 3)
    m(3, 2) = t(2, 3)
    m(3, 3) = t(3, 3)
End Sub

'returns matrix determinant
Public Function GetDeterminantMat4(ByRef m() As Single) As Single
    GetDeterminantMat4 = _
    m(3) * m(6) * m(9) * m(12) - m(2) * m(7) * m(9) * m(12) - m(3) * m(5) * m(10) * m(12) + m(1) * m(7) * m(10) * m(12) + _
    m(2) * m(5) * m(11) * m(12) - m(1) * m(6) * m(11) * m(12) - m(3) * m(6) * m(8) * m(13) + m(2) * m(7) * m(8) * m(13) + _
    m(3) * m(4) * m(10) * m(13) - m(0) * m(7) * m(10) * m(13) - m(2) * m(4) * m(11) * m(13) + m(0) * m(6) * m(11) * m(13) + _
    m(3) * m(5) * m(8) * m(14) - m(1) * m(7) * m(8) * m(14) - m(3) * m(4) * m(9) * m(14) + m(0) * m(7) * m(9) * m(14) + _
    m(1) * m(4) * m(11) * m(14) - m(0) * m(5) * m(11) * m(14) - m(2) * m(5) * m(8) * m(15) + m(1) * m(6) * m(8) * m(15) + _
    m(2) * m(4) * m(9) * m(15) - m(0) * m(6) * m(9) * m(15) - m(1) * m(4) * m(10) * m(15) + m(0) * m(5) * m(10) * m(15)
    
    If GetDeterminantMat4 = 0 Then GetDeterminantMat4 = 1
End Function

'returns matrix inverse
Public Function GetInverseMat4(ByRef m() As Single, ByRef dst() As Single)
    dst(0) = m(6) * m(11) * m(13) - m(7) * m(10) * m(13) + m(7) * m(9) * m(14) - m(5) * m(11) * m(14) - m(6) * m(9) * m(15) + m(5) * m(10) * m(15)
    dst(1) = m(3) * m(10) * m(13) - m(2) * m(11) * m(13) - m(3) * m(9) * m(14) + m(1) * m(11) * m(14) + m(2) * m(9) * m(15) - m(1) * m(10) * m(15)
    dst(2) = m(2) * m(7) * m(13) - m(3) * m(6) * m(13) + m(3) * m(5) * m(14) - m(1) * m(7) * m(14) - m(2) * m(5) * m(15) + m(1) * m(6) * m(15)
    dst(3) = m(3) * m(6) * m(9) - m(2) * m(7) * m(9) - m(3) * m(5) * m(10) + m(1) * m(7) * m(10) + m(2) * m(5) * m(11) - m(1) * m(6) * m(11)
    dst(4) = m(7) * m(10) * m(12) - m(6) * m(11) * m(12) - m(7) * m(8) * m(14) + m(4) * m(11) * m(14) + m(6) * m(8) * m(15) - m(4) * m(10) * m(15)
    dst(5) = m(2) * m(11) * m(12) - m(3) * m(10) * m(12) + m(3) * m(8) * m(14) - m(0) * m(11) * m(14) - m(2) * m(8) * m(15) + m(0) * m(10) * m(15)
    dst(6) = m(3) * m(6) * m(12) - m(2) * m(7) * m(12) - m(3) * m(4) * m(14) + m(0) * m(7) * m(14) + m(2) * m(4) * m(15) - m(0) * m(6) * m(15)
    dst(7) = m(2) * m(7) * m(8) - m(3) * m(6) * m(8) + m(3) * m(4) * m(10) - m(0) * m(7) * m(10) - m(2) * m(4) * m(11) + m(0) * m(6) * m(11)
    dst(8) = m(5) * m(11) * m(12) - m(7) * m(9) * m(12) + m(7) * m(8) * m(13) - m(4) * m(11) * m(13) - m(5) * m(8) * m(15) + m(4) * m(9) * m(15)
    dst(9) = m(3) * m(9) * m(12) - m(1) * m(11) * m(12) - m(3) * m(8) * m(13) + m(0) * m(11) * m(13) + m(1) * m(8) * m(15) - m(0) * m(9) * m(15)
    dst(10) = m(1) * m(7) * m(12) - m(3) * m(5) * m(12) + m(3) * m(4) * m(13) - m(0) * m(7) * m(13) - m(1) * m(4) * m(15) + m(0) * m(5) * m(15)
    dst(11) = m(3) * m(5) * m(8) - m(1) * m(7) * m(8) - m(3) * m(4) * m(9) + m(0) * m(7) * m(9) + m(1) * m(4) * m(11) - m(0) * m(5) * m(11)
    dst(12) = m(6) * m(9) * m(12) - m(5) * m(10) * m(12) - m(6) * m(8) * m(13) + m(4) * m(10) * m(13) + m(5) * m(8) * m(14) - m(4) * m(9) * m(14)
    dst(13) = m(1) * m(10) * m(12) - m(2) * m(9) * m(12) + m(2) * m(8) * m(13) - m(0) * m(10) * m(13) - m(1) * m(8) * m(14) + m(0) * m(9) * m(14)
    dst(14) = m(2) * m(5) * m(12) - m(1) * m(6) * m(12) - m(2) * m(4) * m(13) + m(0) * m(6) * m(13) + m(1) * m(4) * m(14) - m(0) * m(5) * m(14)
    dst(15) = m(1) * m(6) * m(8) - m(2) * m(5) * m(8) + m(2) * m(4) * m(9) - m(0) * m(6) * m(9) - m(1) * m(4) * m(10) + m(0) * m(5) * m(10)
    
    Dim det As Single
    det = GetDeterminantMat4(m)
    If det = 0 Then Exit Function
    
    Dim d As Single
    d = 1 / det
    
    dst(0) = dst(0) * d
    dst(5) = dst(5) * d
    dst(10) = dst(10) * d
    dst(15) = dst(15) * d
End Function

'creates a vector from two points
Public Function Vector3d(ByRef p1 As float3, ByRef p2 As float3) As float3
    Vector3d.x = p1.x - p2.x
    Vector3d.y = p1.y - p2.y
    Vector3d.z = p1.z - p2.z
End Function

'returns magnitude of a vector
Public Function Magnitude(ByRef vector As float3) As Single
    Magnitude = Sqr(vector.x * vector.x + vector.y * vector.y + vector.z * vector.z)
End Function

'rescales vector to the length of one
Public Function Normalize(ByRef vector As float3) As float3
Dim m As Single
    m = Magnitude(vector)
    If m = 0 Then Exit Function 'prevent division by zero
    Normalize.x = vector.x / m
    Normalize.y = vector.y / m
    Normalize.z = vector.z / m
End Function

'generates surface normal from triangle vertices
Public Function GenNormal(ByRef p1 As float3, ByRef p2 As float3, ByRef p3 As float3) As float3
Dim v1 As float3
Dim v2 As float3
Dim n As float3
    v1 = Vector3d(p3, p1)
    v2 = Vector3d(p2, p1)
    n = CrossProduct(v1, v2)
    GenNormal = Normalize(n)
End Function

'...
Public Function TexelToPoint(ByRef v1 As float3, ByRef v2 As float3, ByRef v3 As float3, _
                             ByRef t1 As float2, ByRef t2 As float2, ByRef t3 As float2, _
                             ByRef p As float2) As float3
Dim i As Single
Dim s As Single
Dim t As Single
Dim d As Single
    
    d = ((t2.x - t1.x) * (t3.y - t1.y) - (t2.y - t1.y) * (t3.x - t1.x))
    If d = 0 Then Exit Function 'prevent division by zero
    
    i = 1 / d
    s = i * ((t3.y - t1.y) * (p.x - t1.x) - (t3.x - t1.x) * (p.y - t1.y))
    't = i * ((t2.y - t1.y) * (p.x - t1.x) - (t1.x - t2.x) * (p.y - t1.y))
    t = i * (-(t2.y - t1.y) * (p.x - t1.x) + (t2.x - t1.x) * (p.y - t1.y))
    TexelToPoint.x = v1.x + s * (v2.x - v1.x) + t * (v3.x - v1.x)
    TexelToPoint.y = v1.y + s * (v2.y - v1.y) + t * (v3.y - v1.y)
    TexelToPoint.z = v1.z + s * (v2.z - v1.z) + t * (v3.z - v1.z)
End Function


'implement missing VB math
Public Function ArcSin(ByVal v As Double) As Double
Dim d As Double
    d = Sqr(-v * v + 1)
    If d = 0 Then Exit Function
    ArcSin = Atn(v / d)
End Function
Public Function ArcCos(ByVal v As Double) As Double
Dim d As Double
    d = Sqr(-v * v + 1)
    If d = 0 Then Exit Function
    ArcCos = Atn(-v / d) + 2 * Atn(1)
End Function


'returns the angle between two vector in degrees
Public Function AngleBetweenVectors(ByRef v1 As float3, ByRef v2 As float3) As Single
    AngleBetweenVectors = ArcCos(DotProduct(Normalize(v1), Normalize(v2))) * RADTODEG
End Function


'...
Public Function TrianglePointDistCW(ByRef v1 As float2, ByRef v2 As float2, ByRef v3 As float2, ByRef p As float2) As Single
Dim d1 As Single
Dim d2 As Single
Dim d3 As Single
    d1 = (p.x - v1.x) * (v2.y - v1.y) - (p.y - v1.y) * (v2.x - v1.x) - 1
    d2 = (p.x - v2.x) * (v3.y - v2.y) - (p.y - v2.y) * (v3.x - v2.x) - 1
    d3 = (p.x - v3.x) * (v1.y - v3.y) - (p.y - v3.y) * (v1.x - v3.x) - 1
    'todo
End Function


'...
Public Function PlaneTest(ByRef plane As float4, ByRef point As float3) As Single
Dim n As float3
    n.x = plane.x
    n.y = plane.y
    n.z = plane.z
    PlaneTest = DotProduct(n, point) + plane.w
End Function

'--- triangle-triangle overlap test ------------------------------------------------------------------------------

Public Function ORIENT_2D(ByRef a As float2, ByRef b As float2, ByRef c As float2) As Single
    ORIENT_2D = (a.x - c.x) * (b.y - c.y) - (a.y - c.y) * (b.x - c.x)
End Function


Public Function INTERSECTION_TEST_VERTEX(ByRef p1 As float2, ByRef Q1 As float2, ByRef R1 As float2, _
                                         ByRef p2 As float2, ByRef Q2 As float2, ByRef R2 As float2) As Boolean
Dim ret As Boolean
    If ORIENT_2D(R2, p2, Q1) >= 0 Then
        If ORIENT_2D(R2, Q2, Q1) <= 0 Then
            If ORIENT_2D(p1, p2, Q1) > 0 Then
                If ORIENT_2D(p1, Q2, Q1) <= 0 Then
                    ret = True
                Else
                    ret = False
                End If
            Else
                If ORIENT_2D(p1, p2, R1) >= 0 Then
                    If ORIENT_2D(Q1, R1, p2) >= 0 Then
                        ret = True
                    Else
                        ret = False
                    End If
                Else
                    ret = False
                End If
            End If
        Else
            If ORIENT_2D(p1, Q2, Q1) <= 0 Then
                If ORIENT_2D(R2, Q2, R1) <= 0 Then
                    If ORIENT_2D(Q1, R1, Q2) >= 0 Then
                        ret = True
                    Else
                        ret = False
                    End If
                Else
                    ret = False
                End If
            Else
                ret = False
            End If
        End If
    Else
        If ORIENT_2D(R2, p2, R1) >= 0 Then
            If ORIENT_2D(Q1, R1, R2) >= 0 Then
                If ORIENT_2D(p1, p2, R1) >= 0 Then
                    ret = True
                Else
                    ret = False
                End If
            Else
                If ORIENT_2D(Q1, R1, Q2) >= 0 Then
                    If ORIENT_2D(R2, R1, Q2) >= 0 Then
                        ret = True
                    Else
                        ret = False
                    End If
                Else
                    ret = False
                End If
            End If
        Else
            ret = False
        End If
    End If
    INTERSECTION_TEST_VERTEX = ret
End Function


Public Function INTERSECTION_TEST_EDGE(ByRef p1 As float2, ByRef Q1 As float2, ByRef R1 As float2, _
                                       ByRef p2 As float2, ByRef Q2 As float2, ByRef R2 As float2) As Boolean
Dim ret As Boolean
    If ORIENT_2D(R2, p2, Q1) >= 0 Then
        If ORIENT_2D(p1, p2, Q1) >= 0 Then
            If ORIENT_2D(p1, Q1, R2) >= 0 Then
                ret = True
            Else
                ret = False
            End If
        Else
            If ORIENT_2D(Q1, R1, p2) >= 0 Then
                If ORIENT_2D(R1, p1, p2) >= 0 Then
                    ret = True
                Else
                    ret = False
                End If
            Else
                ret = False
            End If
        End If
    Else
        If ORIENT_2D(R2, p2, R1) >= 0 Then
            If ORIENT_2D(p1, p2, R1) >= 0 Then
                If ORIENT_2D(p1, R1, R2) >= 0 Then
                    ret = True
                Else
                    If ORIENT_2D(Q1, R1, R2) >= 0 Then
                        ret = True
                    Else
                        ret = False
                    End If
                End If
            Else
                ret = False
            End If
        Else
            ret = False
        End If
    End If
    INTERSECTION_TEST_EDGE = ret
End Function


Public Function ccw_tri_tri_intersection_2d(ByRef p1 As float2, ByRef Q1 As float2, ByRef R1 As float2, _
                                            ByRef p2 As float2, ByRef Q2 As float2, ByRef R2 As float2) As Boolean
Dim ret As Boolean
    If ORIENT_2D(p2, Q2, p1) >= 0 Then
        If ORIENT_2D(Q2, R2, p1) >= 0 Then
            If ORIENT_2D(R2, p2, p1) >= 0 Then
                ret = True
            Else
                ret = INTERSECTION_TEST_EDGE(p1, Q1, R1, p2, Q2, R2)
            End If
        Else
            If ORIENT_2D(R2, p2, p1) >= 0 Then
                ret = INTERSECTION_TEST_EDGE(p1, Q1, R1, R2, p2, Q2)
            Else
                ret = INTERSECTION_TEST_VERTEX(p1, Q1, R1, p2, Q2, R2)
            End If
        End If
    Else
         If ORIENT_2D(Q2, R2, p1) >= 0 Then
            If ORIENT_2D(R2, p2, p1) >= 0 Then
                ret = INTERSECTION_TEST_EDGE(p1, Q1, R1, Q2, R2, p2)
            Else
                ret = INTERSECTION_TEST_VERTEX(p1, Q1, R1, Q2, R2, p2)
            End If
         Else
            ret = INTERSECTION_TEST_VERTEX(p1, Q1, R1, R2, p2, Q2)
         End If
    End If
    ccw_tri_tri_intersection_2d = ret
End Function


Public Function TriTriOverlapTest(ByRef p1 As float2, ByRef Q1 As float2, ByRef R1 As float2, _
                                  ByRef p2 As float2, ByRef Q2 As float2, ByRef R2 As float2) As Boolean
Dim ret As Boolean
    If ORIENT_2D(p1, Q1, R1) < 0 Then
        If ORIENT_2D(p2, Q2, R2) < 0 Then
            ret = ccw_tri_tri_intersection_2d(p1, R1, Q1, p2, R2, Q2)
        Else
            ret = ccw_tri_tri_intersection_2d(p1, R1, Q1, p2, Q2, R2)
        End If
    Else
        If ORIENT_2D(p2, Q2, R2) < 0 Then
            ret = ccw_tri_tri_intersection_2d(p1, Q1, R1, p2, R2, Q2)
        Else
            ret = ccw_tri_tri_intersection_2d(p1, Q1, R1, p2, Q2, R2)
        End If
    End If
    TriTriOverlapTest = ret
End Function

'--- power of two -----------------------------------

Public Function PowTwo(ByVal n As Long) As Long
Dim i As Long
Dim v As Long
    For i = 1 To n
        v = v * 2
    Next i
    PowTwo = v
End Function

Public Function DivTwo(ByVal n As Long) As Long
Dim i As Long
Dim v As Long
    i = 0
    v = n
    Do
        v = v / 2
        If v > 0 Then
            i = i + 1
        Else
            DivTwo = i
            Exit Function
        End If
    Loop
End Function

'-------------------------------------------------------------------------------------------------------------------------

'returns 1 if point lies inside triangle edges, 2 if within edge margin, 0 if outside
Public Function InsideTriangle(t1 As float2, t2 As float2, t3 As float2, p As float2, ByVal margin As Single) As Long
    
    'check if point is inside triangle
    If TriangleTest(t1, t2, t3, p) Then
        InsideTriangle = 1
        Exit Function
    End If
    
    'outside triangle, but see if distance is within edge margin
    
    If margin > 0 Then
        
        't1-t2
        If PointDistToSegment(p, t1, t2) < margin Then
            InsideTriangle = 2
            Exit Function
        End If
        
        't2-t3
        If PointDistToSegment(p, t2, t3) < margin Then
            InsideTriangle = 2
            Exit Function
        End If
        
        't3-t1
        If PointDistToSegment(p, t3, t1) < margin Then
            InsideTriangle = 2
            Exit Function
        End If
        
    End If
    
    'definately outside triangle
    InsideTriangle = 0
End Function


'returns closest position to point on triangle
Public Function ClosestPointOnTriangle(t1 As float2, t2 As float2, t3 As float2, p As float2) As float2
    
    'check if point is inside triangle
    If TriangleTest(t1, t2, t3, p) Then
        ClosestPointOnTriangle = p
        Exit Function
    End If
    
    'outside triangle, return closest point on edges
    
    Dim cp As float2
    Dim d As Single
    
    Dim tcp As float2
    Dim td As Single
    
    't1-t2
    cp = ClosestPointOnLine(t1, t2, p)
    d = Distance2d(cp, p)
    
    't2-t3
    tcp = ClosestPointOnLine(t2, t3, p)
    td = Distance2d(tcp, p)
    If td < d Then
        cp = tcp
        d = td
    End If
    
    't3-t1
    tcp = ClosestPointOnLine(t3, t1, p)
    td = Distance2d(tcp, p)
    If td < d Then
        cp = tcp
        d = td
    End If
    
    ClosestPointOnTriangle = cp
End Function


'returns true if point lies inside triangle (both CW and CCW winding)
Public Function TriangleTest(ByRef v1 As float2, ByRef v2 As float2, ByRef v3 As float2, ByRef p As float2) As Boolean
    TriangleTest = TriangleTestCW(v1, v2, v3, p) Or TriangleTestCW(v3, v2, v1, p)
End Function


'returns true if point lies inside triangle (clockwise vertex order)
Public Function TriangleTestCW(ByRef v1 As float2, ByRef v2 As float2, ByRef v3 As float2, ByRef p As float2) As Boolean
    If (p.x - v1.x) * (v2.y - v1.y) - (p.y - v1.y) * (v2.x - v1.x) > 0 Then Exit Function
    If (p.x - v2.x) * (v3.y - v2.y) - (p.y - v2.y) * (v3.x - v2.x) > 0 Then Exit Function
    If (p.x - v3.x) * (v1.y - v3.y) - (p.y - v3.y) * (v1.x - v3.x) > 0 Then Exit Function
    TriangleTestCW = True
End Function


'returns distance to line segment between two points
Public Function PointDistToSegment(p As float2, v1 As float2, v2 As float2) As Single
Dim v As float2
Dim w As float2
    
    v.x = v2.x - v1.x
    v.y = v2.y - v1.y
    w.x = p.x - v1.x
    w.y = p.y - v1.y
    
    Dim c1 As Single
    c1 = DotProduct2d(w, v)
    If c1 <= 0 Then
        PointDistToSegment = Distance2d(p, v1)
        Exit Function
    End If

    Dim c2 As Single
    c2 = DotProduct2d(v, v)
    If c2 <= c1 Then
        PointDistToSegment = Distance2d(p, v2)
        Exit Function
    End If
        
    Dim b As Single
    If c2 = 0 Then
        b = 0
    Else
        b = c1 / c2
    End If
     
    Dim pb As float2
    pb.x = v1.x + b * v.x
    pb.y = v1.y + b * v.y
    
    PointDistToSegment = Distance2d(p, pb)
End Function


'returns distance between two 2d points
Public Function Distance2d(a As float2, b As float2) As Single
    Distance2d = Sqr(((a.x - b.x) * (a.x - b.x)) + ((a.y - b.y) * (a.y - b.y)))
End Function


'2d dot product
Public Function DotProduct2d(ByRef v1 As float2, ByRef v2 As float2) As Single
    DotProduct2d = (v1.x * v2.x + v1.y * v2.y)
End Function


'returns closest point on infinite line
Public Function PointDistToLine(p As float2, v1 As float2, v2 As float2) As Single
Dim v As float2
Dim w As float2
    
    v.x = v2.x - v1.x
    v.y = v2.y - v1.y
    
    w.x = p.x - v1.x
    w.y = p.y - v1.y
    
Dim c1 As Single
Dim c2 As Single
    c1 = DotProduct2d(w, v)
    c2 = DotProduct2d(v, v)
    
    Dim b As Single
    If c2 = 0 Then
        b = 0
    Else
        b = c1 / c2
    End If
    
    Dim pb As float2
    pb.x = v1.x + b * v.x
    pb.x = v1.x + b * v.x
    
    PointDistToLine = Distance2d(p, pb)
End Function

'''''''''''''''''''

Public Function ClosestPointOnLine(ByRef a As float2, ByRef b As float2, ByRef p As float2) As float2
    Dim ap As float2
    Dim ab As float2
    
    ap.x = p.x - a.x
    ap.y = p.y - a.y
    
    ab.x = b.x - a.x
    ab.y = b.y - a.y
    
    Dim ab2 As Single
    Dim ap_ab As Single
    Dim t As Single
    ab2 = ab.x * ab.x + ab.y * ab.y
    ap_ab = ap.x * ab.x + ap.y * ab.y
    
    If ab2 = 0 Then
        t = ap_ab
    Else
        t = ap_ab / ab2
    End If
    
    If t < 0 Then t = 0
    If t > 1 Then t = 1
    
    Dim closest As float2
    closest.x = a.x + ab.x * t
    closest.y = a.y + ab.y * t
    
    ClosestPointOnLine = closest
End Function

'--- quaternions -------------------------------------------------------------------

Public Sub QuatIdentity(ByRef q As quat)
    q.x = 0
    q.y = 0
    q.z = 0
    q.w = 1
End Sub

Public Function QuatRot(ByRef r As quat, ByRef vec As float3) As float3
Dim q As quat
    q.x = (vec.x * r.w) + (vec.z * r.y) - (vec.y * r.z)
    q.y = (vec.y * r.w) + (vec.x * r.z) - (vec.z * r.x)
    q.z = (vec.z * r.w) + (vec.y * r.x) - (vec.x * r.y)
    q.w = (vec.x * r.x) + (vec.y * r.y) + (vec.z * r.z)
    QuatRot.x = (r.w * q.x) + (r.x * q.w) + (r.y * q.z) - (r.z * q.y)
    QuatRot.y = (r.w * q.y) + (r.y * q.w) + (r.z * q.x) - (r.x * q.z)
    QuatRot.z = (r.w * q.z) + (r.z * q.w) + (r.x * q.y) - (r.y * q.x)
End Function

Public Function QuatInv(ByRef q As quat) As quat
Dim s As Single
    s = (1 / ((q.x * q.x) + (q.y * q.y) + (q.z * q.z) + (q.w * q.w)))
    QuatInv.x = q.x * -s
    QuatInv.y = q.y * -s
    QuatInv.z = q.z * -s
    QuatInv.w = q.w * s
End Function

Public Function QuatMul(ByRef a As quat, ByRef b As quat) As quat
     QuatMul.x = (a.x * b.w) + (a.y * b.z) - (a.z * b.y) + (a.w * b.x)
    QuatMul.y = (-a.x * b.z) + (a.y * b.w) + (a.z * b.x) + (a.w * b.y)
     QuatMul.z = (a.x * b.y) - (a.y * b.x) + (a.z * b.w) + (a.w * b.z)
    QuatMul.w = (-a.x * b.x) - (a.y * b.y) - (a.z * b.z) + (a.w * b.w)
End Function

Public Function QuatAdd(ByRef a As quat, ByRef b As quat) As quat
    QuatAdd.x = a.x + b.x
    QuatAdd.y = a.y + b.y
    QuatAdd.z = a.z + b.z
    QuatAdd.w = a.w + b.w
End Function

Public Function QuatMagnitude(ByRef q As quat) As Single
    QuatMagnitude = Sqr((q.x * q.x) + (q.y * q.y) + (q.z * q.z) + (q.w * q.w))
End Function

Public Sub QuatNormalize(ByRef q As quat)
    Dim m As Single
    m = QuatMagnitude(q)
    If m = 0 Then
        QuatIdentity q
        Exit Sub
    End If
    q.x = q.x / m
    q.y = q.y / m
    q.z = q.z / m
    q.w = q.w / m
End Sub
