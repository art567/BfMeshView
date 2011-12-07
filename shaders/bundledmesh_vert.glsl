#version 120

uniform float hasBump;
uniform float hasWreck;
uniform float hasXnimatedUV;

uniform mat4 nodetransform[40];
uniform vec3 eyeposworld;
//uniform vec3 sunvecworld;

varying vec2 uv;
varying vec3 norm;
varying vec3 eyesurfvec;         // eye to surface vector
varying vec3 sunvec;
varying vec4 boneinfo;

void main()
{
 //// temp: pass as uniform!
 //vec3 sunvecworld = normalize(vec3(0.5, -0.5, 0.5));
 vec3 sunvecworld = -eyeposworld ;
 //// temp
 
 // bone id
 int nodeid = int(gl_Color.r*255.0);
 
 // transform vertex
 vec4 vert = nodetransform[ nodeid ] * gl_Vertex;
 
 // UV0
 uv = gl_MultiTexCoord0.st;
 
 if (hasXnimatedUV > 0.5) {
  uv += gl_MultiTexCoord1.st * vec2(0.5,1.0);
 }
 
 // normal
 norm = gl_Normal;
 
 // eye to surface vector in world space
 eyesurfvec = eyeposworld - vert.xyz;
 
 // transform sunvec
 sunvec = (vec4(sunvecworld,0.0) * nodetransform[ nodeid ]).xyz;
 
 // tangent
 if (hasBump > 0.5) {
  
  // compute tangents
  vec3 tan1 = gl_MultiTexCoord5.xyz;
  vec3 tan2 = cross(norm,-tan1);
  
  // create tangent space rotation matrix
  mat3 rotmat = mat3(tan1,tan2,norm);
  
  // rotate local space sun vector to tangent space
  sunvec = sunvec * rotmat;
  
  // rotate local eye-to-surface to tangent space
  eyesurfvec = eyesurfvec * rotmat;
 }
 
 boneinfo = gl_Color;
 
 // vertex position
 //gl_Position = gl_ModelViewProjectionMatrix * vert;
 vec4 v = gl_ModelViewMatrix * vert;
 v.w -= 0.01;
 gl_Position = gl_ProjectionMatrix * v;
}
