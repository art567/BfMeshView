#version 120

uniform int hasBump;
uniform int hasWreck;
uniform int hasAnimatedUV;

uniform mat4 nodetransform[40];
uniform vec3 eyeposworld;
uniform vec3 eyevecworld;
//uniform vec3 sunvecworld;

varying vec2 uv;
varying vec3 norm;
varying vec3 eyesurfvec;         // eye to surface vector
varying vec3 sunvec;
varying vec4 boneinfo;
varying vec3 eyepos;
varying vec3 fragpos;
varying float debug;

void main()
{
 //// temp: pass as uniform!
 //vec3 sunvecworld = normalize(vec3(0.5, -0.5, 0.5));
 vec3 sunvecworld = eyevecworld;
 //// temp
 
 // bone id
 int nodeid = int(gl_Color.r*255.0);
 
 // transform vertex
 vec4 vert = nodetransform[ nodeid ] * gl_Vertex;
 
 // UV0
 uv = gl_MultiTexCoord0.st;
 
 if (hasAnimatedUV > 0) {
  uv += gl_MultiTexCoord1.st * vec2(0.5,1.0);
 }
 
 // normal
 norm = (nodetransform[ nodeid ] * vec4(gl_Normal,0.0)).xyz;
 norm = gl_Normal;
 
 // transform eye position to node space
 vec3 eyeposlocal = (vec4(eyeposworld,0.0) * nodetransform[ nodeid ]).xyz;
 
 // transform sunvec
 sunvec = (vec4(sunvecworld,0.0) * nodetransform[ nodeid ]).xyz;
 
 // eye to surface vector in world space
 eyesurfvec = eyeposlocal - vert.xyz;
 
 // tangent
 if (hasBump > 0) {
  
  // compute tangents
  vec3 tan1 = gl_MultiTexCoord5.xyz;
  vec3 tan2 = cross(gl_Normal,-tan1) * gl_MultiTexCoord5.w;
  
  // create tangent space rotation matrix
  mat3 rotmat = mat3(tan1,tan2,gl_Normal);
  
  // rotate local space sun vector to tangent space
  sunvec = sunvec * rotmat;
  
  // rotate local eye-to-surface to tangent space
  eyesurfvec = eyesurfvec * rotmat;
 }
 
 boneinfo = gl_Color;
 
 eyepos = eyeposworld;
 fragpos = vert.xyz;
 debug = gl_MultiTexCoord5.w;
 
 // vertex position
 gl_Position = gl_ModelViewProjectionMatrix * vert;
}
