#version 120

uniform float timephase;
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
 //gl_Position = gl_ModelViewProjectionMatrix * gl_Vertex;
 
 vec4 v = gl_Vertex;
 //v.y += sin(timephase);
 
 
 float GlobalTime = timephase * 6.0;
 float WindSpeed = 5.0;
 float LEAF_MOVEMENT = 1024.0;
 float ObjRadius = 1.0;
 
 //float fh2wAmount = ObjRadius + min(v.y,10.0);
	//v.xyz += sin((GlobalTime / fh2wAmount) * WindSpeed) * fh2wAmount * fh2wAmount / LEAF_MOVEMENT;
 
 // BF2
 //v.xyz +=  sin((GlobalTime / (ObjRadius + v.y)) * WindSpeed) * (ObjRadius + v.y) * (ObjRadius + v.y) / LEAF_MOVEMENT;
 
 float asd = ObjRadius + min(v.y,5.0);
 v.xyz += sin((GlobalTime / (ObjRadius+v.y)*0.5) * WindSpeed) * (asd) * (asd) / LEAF_MOVEMENT;
 
 
 gl_Position = gl_ModelViewProjectionMatrix * v;
}
