#version 120

uniform sampler2D texture0;

uniform int showDiffuse;
uniform int showLighting;

//uniform vec3 sunambient;
//uniform vec3 sundiffuse;

varying vec2 uv;
varying vec3 norm;
varying vec3 sunvec;

varying vec3 eyepos;
varying vec3 fragpos;

void main()
{
 //// temp: pass as uniforms
 vec3 sunambient = vec3(0.3,0.3,0.3);
 vec3 sundiffuse = vec3(0.7,0.7,0.7);
 vec3 sunspecular = sundiffuse;
 //// temp
 
 // base
 vec4 frag = vec4(1.0, 1.0, 1.0, 1.0);
 
 // textures
 vec4 base = texture2D(texture0, uv);
 
 if (showDiffuse > 0) {
  frag.rgb *= base.rgb;
 }
 
 // alpha
 frag.a *= base.a*2.0;
 
 // normal
 vec3 n = normalize(norm);
 
 // lighting
 if (showLighting > 0) {
  float NdotL = dot(n,normalize(-sunvec));
  frag.rgb *= sunambient.rgb + sundiffuse.rgb * max(NdotL,0.0);
 }
 
 // output
 gl_FragColor = frag;
 //gl_FragColor.rgb = n * 0.5 + 0.5;
}
