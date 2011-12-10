#version 120

uniform sampler2D texture0; // base
uniform sampler2D texture1; // detail

uniform int showDiffuse;
uniform int showLighting;

//uniform vec3 sunambient;
//uniform vec3 sundiffuse;

varying vec2 uv0;
varying vec2 uv1;
varying vec3 norm;
varying vec3 sunvec;

void main()
{
 //// temp: pass as uniforms
 vec3 sunambient = vec3(0.3,0.3,0.3);
 vec3 sundiffuse = vec3(0.7,0.7,0.7);
 vec3 sunspecular = sundiffuse;
 //// temp
 
 // base
 vec4 frag = vec4(1.0, 1.0, 1.0, 1.0);
 vec3 spec = vec3(1.0, 1.0, 1.0);
 
 // textures
 vec4 basemap    = texture2D(texture0, uv0);
 vec4 detailmap  = texture2D(texture1, uv1);
 
 // diffuse
 if (showDiffuse > 0) {
  frag *= basemap*2.0;
  frag *= detailmap;
 } else {
  frag.rgb *= 0.75;
 }
 
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
