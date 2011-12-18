Attribute VB_Name = "BF2_MeshStruct"
Option Explicit

'internal struct, we use this for our approximate shading
Public Type mat_layer
    texmapid As Long
    texcoff As Long
    blend As Boolean
    blendsrc As GLuint
    blenddst As GLuint
    alphaTest As Boolean
    alpharef As Single
    depthfunc As GLenum
    depthWrite As GLboolean
    twosided As Boolean
    lighting As Boolean
End Type

'---------------------------------

'bf2 mesh file header
Public Type bf2head '20 bytes
    u1 As Long          '0
    version As Long     '10 for most bundledmesh, 6 for some bundledmesh, 11 for staticmesh
    u3 As Long          '0
    u4 As Long          '0
    u5 As Long          '0
End Type


'vertex attribute table entry
Public Type bf2vertattrib '8 bytes
    flag As Integer         'some sort of boolean flag (if true the below field are to be ignored?)
    offset As Integer       'offset from vertex data start
    vartype As Integer      'attribute type (vec2, vec3 etc)
    usage As Integer        'usage ID (vertex, texcoord etc)
    
    'Note: "usage" field correspond to the definition in DX SDK "Include\d3d9types.h"
    '      It looks like DICE extended these for additional UV channels, these
    '      constants are much larger so they don't conflict with other DX enums.
End Type


'bone structure
Public Type bf2bone  '68 bytes
    id As Long        'bone ID (4 bytes)
    matrix As matrix4 'inverse bone matrix (64 bytes)
    
    'internal
    skinmat As matrix4  'world space deformed skin transform
End Type

'rig structure
Public Type bf2rig
    bonenum As Long
    bone() As bf2bone
End Type


'lod drawcall
Public Type bf2mat
    alphamode As Long     '0=opaque, 1=blend, 2=alphatest
    fxfile As String      'shader filename string
    technique As String   'technique name
    
    'texture map filenames
    mapnum As Long
    map() As String
    
    'geometry info
    vstart As Long      'vertex start offset
    istart As Long      'index start offset
    inum As Long        'number of indices??
    vnum As Long        'number of vertices??
    
    'unknown
    u4 As Long          '0
    u5 As Long          '0
    
    'per-material bounds (staticmesh only)
    mmin As float3
    mmax As float3
    
    ''''internal
    texmapid() As Long      'texmap() index
    mapuvid() As Long       'UV index for each map
    layernum As Long
    layer(1 To 4) As mat_layer
    glslprog As Long
    hasBump As Boolean
    hasWreck As Boolean
    hasAnimatedUV As Boolean
    hasBumpAlpha As Boolean
    hasDirt As Boolean
    hasCrack As Boolean
    hasCrackN As Boolean
    hasDetailN As Boolean
    hasEnvMap As Boolean
    alphaTest As Single
    twosided As Boolean
End Type


'bf2 lod
Public Type bf2lod
    
    'bounds
    min As float3
    max As float3
    pivot As float3 'not sure this is really a pivot (only on version<=6)
    
    'skinning matrices (skinnedmesh only)
    rignum As Long  'this corresponds to matnum
    rig() As bf2rig
    
    'nodes (staticmesh and bundledmesh only)
    nodenum As Long
    node() As matrix4
    
    'material groups
    matnum As Long
    mat() As bf2mat
    
    ''''internal
    polycount As Long
End Type


'bf2 geom
Public Type bf2geom
    lodnum As Long
    lod() As bf2lod
End Type


'bf2 BundledMesh vertex weight (helper structure, memcopy float to this)
Public Type bf2vw
    b1 As Byte 'bone 1 index
    b2 As Byte 'bone 2 index
    w1 As Byte 'weight for bone 1
    w2 As Byte 'weight for bone 2
End Type

'bf2 SkinnedMesh vertex weight (helper structure, memcopy float to this)
Public Type bf2skinweight
    w As Single
    b1 As Byte
    b2 As Byte
    b3 As Byte
    b4 As Byte
End Type

'bf2 vertex info (helper structure generated after load time
Public Type bf2vertinfo
    geom As Byte
    lod As Byte
    mat As Byte
    sel As Byte 'unused
End Type


'file structure
Public Type bf2mesh
    
    'header
    head As bf2head
    
    'unknown
    u1 As Byte 'always 0?
    
    'geoms
    geomnum As Long
    geom() As bf2geom
    
    'vertex attribute table
    vertattribnum As Long
    vertattrib() As bf2vertattrib
    
    'vertices
    vertformat As Long 'always 4?  (e.g. GL_FLOAT)
    vertstride As Long
    vertnum As Long
    vert() As Single
    
    'indices
    indexnum As Long
    Index() As Integer
    
    'unknown
    u2 As Long 'always 8?
    
    ''''internal
    filename As String       'current loaded mesh file
    fileext As String        'filename extension
    isSkinnedMesh As Boolean 'true if file extension is "skinnedmesh"
    isBundledMesh As Boolean 'true if file extension is "bundledmesh"
    isBFP4F As Boolean       'true if file is inside BFP4F directory
    loadok As Boolean        'mesh loaded properly
    drawok As Boolean        'mesh rendered properly
    xstride As Long          'vertstride / 4
    uvnum As Long            'number of detected uv channels
    
    vertinfo() As bf2vertinfo
    vertsel() As Byte        'vertex selection flags
    vertflag() As Byte       'per vertex flag for various things
    
    hasSkinVerts As Boolean  'deformed vertices flag
    skinvert() As float3     'deformed vertices
    skinnorm() As float3     'deformed normals
End Type


'--- helper functions -----------------------------------------------------------------------------------


'returns vertex buffer offset for normal attribute
Private Function BF2MeshGetAttribOffset(ByVal usageID As Long) As Long
    Dim i As Long
    With vmesh
        For i = 0 To .vertattribnum - 1
            If .vertattrib(i).usage = usageID Then
                BF2MeshGetAttribOffset = .vertattrib(i).offset / 4
                Exit Function
            End If
        Next i
    End With
    BF2MeshGetAttribOffset = -1
End Function


'returns vertex buffer offset for UV attribute
Public Function BF2MeshGetTexcOffset(ByVal uvchan As Long) As Long
    BF2MeshGetTexcOffset = -1
    If uvchan = 0 Then BF2MeshGetTexcOffset = BF2MeshGetAttribOffset(5)
    If uvchan = 1 Then BF2MeshGetTexcOffset = BF2MeshGetAttribOffset(261)
    If uvchan = 2 Then BF2MeshGetTexcOffset = BF2MeshGetAttribOffset(517)
    If uvchan = 3 Then BF2MeshGetTexcOffset = BF2MeshGetAttribOffset(773)
    If uvchan = 4 Then BF2MeshGetTexcOffset = BF2MeshGetAttribOffset(1029)
End Function


'returns vertex buffer offset for normal vector
Public Function BF2MeshGetNormOffset() As Long
    BF2MeshGetNormOffset = BF2MeshGetAttribOffset(3)
End Function


'returns vertex buffer offset for tangent vector
Public Function BF2MeshGetTangOffset() As Long
    BF2MeshGetTangOffset = BF2MeshGetAttribOffset(6)
End Function


'returns vertex buffer offset for weight/bone index attributes (start of 4 x 1-byte block)
Public Function BF2MeshGetWeightOffset() As Long
    BF2MeshGetWeightOffset = BF2MeshGetAttribOffset(2)
End Function

