Attribute VB_Name = "a_GLEXT"
Option Explicit


'GLext
Public Declare Function glextInit Lib "glext.dll" () As Boolean


'GL_ARB_texture_compression
Public Const GL_COMPRESSED_RGB_S3TC_DXT1_EXT = 33776  'todo: does our code support this?
Public Const GL_COMPRESSED_RGBA_S3TC_DXT1_EXT = 33777
Public Const GL_COMPRESSED_RGBA_S3TC_DXT3_EXT = 33778
Public Const GL_COMPRESSED_RGBA_S3TC_DXT5_EXT = 33779
Public Declare Sub glCompressedTexImage2D Lib "glext.dll" (ByVal target As GLenum, _
                                                           ByVal level As GLint, ByVal internalformat As GLint, _
                                                           ByVal width As GLsizei, ByVal height As GLsizei, _
                                                           ByVal border As GLint, _
                                                           ByVal imagesize As GLsizei, ByVal data As Long)

'GL_ARB_multitexture
Public Const GL_TEXTURE0 = 33984
Public Const GL_TEXTURE1 = 33985
Public Const GL_TEXTURE2 = 33986
Public Const GL_TEXTURE3 = 33987
Public Const GL_TEXTURE4 = 33988
Public Const GL_TEXTURE5 = 33989
Public Const GL_TEXTURE6 = 33990
Public Const GL_TEXTURE7 = 33991
Public Const GL_TEXTURE8 = 33992
Public Const GL_ACTIVE_TEXTURE = 34016
Public Const GL_CLIENT_ACTIVE_TEXTURE = 34017
Public Const GL_MAX_TEXTURE_UNITS = 34018
Public Declare Sub glActiveTexture Lib "glext32.dll" (ByVal texture As GLenum)
Public Declare Sub glClientActiveTexture Lib "glext32.dll" (ByVal texture As GLenum)
Public Declare Sub glMultiTexCoord2f Lib "glext32.dll" (ByVal target As GLenum, ByVal s As GLfloat, ByVal t As GLfloat)

'GL_ARB_shader_objects
Public Const GL_PROGRAM_OBJECT = 35648
Public Const GL_SHADER_OBJECT = 35656
Public Const GL_OBJECT_TYPE = 35662
Public Const GL_OBJECT_SUBTYPE = 35663
Public Const GL_FLOAT_VEC2 = 35664
Public Const GL_FLOAT_VEC3 = 35665
Public Const GL_FLOAT_VEC4 = 35666
Public Const GL_INT_VEC2 = 35667
Public Const GL_INT_VEC3 = 35668
Public Const GL_INT_VEC4 = 35669
Public Const GL_BOOL = 35670
Public Const GL_BOOL_VEC2 = 35671
Public Const GL_BOOL_VEC3 = 35672
Public Const GL_BOOL_VEC4 = 35673
Public Const GL_FLOAT_MAT2 = 35674
Public Const GL_FLOAT_MAT3 = 35675
Public Const GL_FLOAT_MAT4 = 35676
Public Const GL_SAMPLER_1D = 35677
Public Const GL_SAMPLER_2D = 35678
Public Const GL_SAMPLER_3D = 35679
Public Const GL_SAMPLER_CUBE = 35680
Public Const GL_SAMPLER_1D_SHADOW = 35681
Public Const GL_SAMPLER_2D_SHADOW = 35682
Public Const GL_SAMPLER_2D_RECT = 35683
Public Const GL_SAMPLER_2D_RECT_SHADOW = 35684
Public Const GL_OBJECT_DELETE_STATUS = 35712
Public Const GL_OBJECT_COMPILE_STATUS = 35713
Public Const GL_OBJECT_LINK_STATUS = 35714
Public Const GL_OBJECT_VALIDATE_STATUS = 35715
Public Const GL_OBJECT_INFO_LOG_LENGTH = 35716
Public Const GL_OBJECT_ATTACHED_OBJECTS = 35717
Public Const GL_OBJECT_ACTIVE_UNIFORMS = 35718
Public Const GL_OBJECT_ACTIVE_UNIFORM_MAX_LENGTH = 35719
Public Const GL_OBJECT_SHADER_SOURCE_LENGTH = 35720

'GL_ARB_vertex_shader
Public Const GL_VERTEX_SHADER = 35633
Public Const GL_MAX_VERTEX_UNIFORM_COMPONENTS = 35658
Public Const GL_MAX_VARYING_FLOATS = 35659
Public Const GL_MAX_VERTEX_TEXTURE_IMAGE_UNITS = 35660
Public Const GL_MAX_COMBINED_TEXTURE_IMAGE_UNITS = 35661
Public Const GL_OBJECT_ACTIVE_ATTRIBUTES = 35721
Public Const GL_OBJECT_ACTIVE_ATTRIBUTE_MAX_LENGTH = 35722

'GL_ARB_fragment_shader
Public Const GL_FRAGMENT_SHADER = 35632
Public Const GL_MAX_FRAGMENT_UNIFORM_COMPONENTS = 35657
Public Const GL_FRAGMENT_SHADER_DERIVATIVE_HINT = 35723

'GL_ARB_texture_border_clamp
Public Const GL_CLAMP_TO_BORDER = 33069

'misc
Public Const GL_COMBINE_ARB = 34160
Public Const GL_RGB_SCALE_ARB = 34163
Public Const GL_TEXTURE_MAX_ANISOTROPY_EXT = 34046
Public Const GL_MAX_TEXTURE_MAX_ANISOTROPY_EXT = 34047

