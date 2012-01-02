#version 120

uniform int hasBump;

uniform vec3 eyeposworld;
uniform vec3 eyevecworld;
//uniform vec3 sunvecworld;

varying vec2 uv0;
varying vec2 uv1;
varying vec2 uv2;
varying vec2 uv3;
varying vec2 uv4;
varying vec3 norm;
varying vec3 eyesurfvec;         // eye to surface vector
varying vec3 sunvec;
varying vec3 eyepos;
varying vec3 fragpos;

void main()
{
 //// temp: pass as uniform!
 //vec3 sunvecworld = normalize(vec3(0.5, -0.5, 0.5));
 vec3 sunvecworld = eyevecworld;
 //// temp
 
 // UVs
 uv0 = gl_MultiTexCoord0.st;
 uv1 = gl_MultiTexCoord1.st;
 uv2 = gl_MultiTexCoord2.st;
 uv3 = gl_MultiTexCoord3.st;
 uv4 = gl_MultiTexCoord4.st;
 
 // normal
 norm = gl_Normal;
 
 // transform sunvec
 sunvec = sunvecworld;
 
 // eye to surface vector in world space
 eyesurfvec = eyeposworld - gl_Vertex.xyz;
 
 eyepos = eyeposworld;
 fragpos = gl_Vertex.xyz;
 
 // tangent
 if (hasBump > 0) {
  
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
 
 // vertex position
 gl_Position = gl_ModelViewProjectionMatrix * gl_Vertex;
}
