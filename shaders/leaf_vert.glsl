#version 120

uniform vec3 eyeposworld;
uniform vec3 eyevecworld;
//uniform vec3 sunvecworld;

varying vec2 uv;
varying vec3 norm;
varying vec3 sunvec;

void main()
{
 //// temp: pass as uniform!
 vec3 sunvecworld = normalize(vec3(0.5, -0.5, 0.5));
 //vec3 sunvecworld = eyevecworld;
 //// temp
 
 // UVs
 uv = gl_MultiTexCoord0.st; 
 
 // normal
 norm = gl_Vertex.xyz;// gl_Normal;
 
 // transform sunvec
 sunvec = sunvecworld;
 
 // vertex position
 gl_Position = gl_ModelViewProjectionMatrix * gl_Vertex;
}
