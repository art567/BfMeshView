#version 120

uniform sampler2D texture0; // base
uniform sampler2D texture1; // detail
uniform sampler2D texture2; // dirt
uniform sampler2D texture3; // crack
uniform sampler2D texture4; // detailN
uniform sampler2D texture5; // crackN

//uniform int hasBump;
uniform int hasAlpha;
uniform int hasDirt;
uniform int hasCrack;
uniform int hasCrackN;
uniform int hasDetailN;
uniform int showDiffuse;
uniform int showLighting;

//uniform vec3 sunambient;
//uniform vec3 sundiffuse;

varying vec2 uv0;
varying vec2 uv1;
varying vec2 uv2;
varying vec2 uv3;
varying vec2 uv4;
varying vec3 norm;
varying vec3 sunvec;
varying vec3 eyesurfvec;         // eye to surface vector

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
 vec3 spec = vec3(1.0, 1.0, 1.0);
 
 // textures
 vec4 basemap    = texture2D(texture0, uv0);
 vec4 detailmap  = texture2D(texture1, uv1);
 vec4 dirtmap    = texture2D(texture2, uv2);
 vec4 crackmap   = texture2D(texture3, uv3);
 vec4 detailmapN;
 vec4 crackmapN;
 
 if (hasDirt > 0) {
  if (hasCrack > 0) {
   detailmapN = texture2D(texture4, uv1);
   crackmapN = texture2D(texture5, uv3);
  } else {
   detailmapN = texture2D(texture3, uv1);
  }
 } else {
  if (hasCrack > 0) {
   detailmapN = texture2D(texture3, uv1);
   crackmap = texture2D(texture2, uv2);
   crackmapN = texture2D(texture4, uv2);
  } else {
   detailmapN = texture2D(texture2, uv1);
  }
 }
 
 // diffuse
 if (showDiffuse > 0) {
  frag *= basemap;
  frag *= detailmap;
 } else {
  frag.rgb *= 0.75;
 }
 
 // alpha
 if (hasAlpha > 0) {
  frag.a = detailmap.a;
 }
 
 // crack
 if (hasCrack > 0) {
  if (showDiffuse > 0) frag.rgb = mix(frag.rgb, crackmap.rgb, crackmap.a);
 }
 
 // dirt
 if (hasDirt > 0) {
  if (showDiffuse > 0) frag.rgb *= dirtmap.rgb;
  spec *= dirtmap.r*dirtmap.g*dirtmap.b;
 }
 
 // specular
 if (hasAlpha > 0) {
  if (hasDetailN > 0) {
   spec *= detailmapN.a;
  }
 } else {
  spec *= detailmap.a;
 }
 
 // normal
 vec3 n;
 if (hasDetailN > 0) {
  n = detailmapN.rgb * 2.0 - 1.0; // detail normal
  if (hasCrack > 0) {
   n = mix(n, crackmapN.rgb * 2.0 - 1.0, crackmap.a); // crack normal
  }
 } else {
  n = vec3(0.0, 0.0, 1.0); // vertex normal
 }
 n = normalize(n);
 
 // lighting
 if (showLighting > 0) {
  float NdotL = dot(n,normalize(-sunvec));
  frag.rgb *= sunambient.rgb + sundiffuse.rgb * max(NdotL,0.0);
  
  // specular
  if (NdotL > 0.0) {
   
   // half vector
   vec3 hv = normalize( -sunvec + eyesurfvec );
   
   // compute specular amount
   float NdotHV = max(dot(n,hv),0.0);
   spec *= pow(NdotHV,100.0);
   
   // apply specular
   frag.rgb += sunspecular * spec;
  }
 }
 
 // output
 gl_FragColor = frag;
 //gl_FragColor.rgb = n * 0.5 + 0.5;
}
