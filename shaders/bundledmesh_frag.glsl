#version 120

uniform sampler2D texture0; // diffuse
uniform sampler2D texture1; // normal
//uniform sampler2D texture2; // SpecularLUT
uniform sampler2D texture3; // wreck
uniform samplerCube envmap; // envmap

uniform int hasBump;
uniform int hasWreck;
uniform int hasAlpha;
uniform int hasEnvMap;
uniform int hasBumpAlpha;
uniform int showDiffuse;
uniform int showLighting;

//uniform vec3 sunambient;
//uniform vec3 sundiffuse;

varying vec2 uv;
varying vec3 norm;
varying vec3 sunvec;
varying vec3 eyesurfvec;         // eye to surface vector
varying vec4 boneinfo;

varying vec3 eyepos;
varying vec3 fragpos;
varying float debug;

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
 vec4 colormap = texture2D(texture0, uv);
 vec4 normalmap = texture2D(texture1, uv);
 vec4 wreckmap = texture2D(texture3, uv);
 
 // vertex normal
 vec3 vertNorm = normalize(norm);
 
 // diffuse
 if (showDiffuse > 0) {
  frag *= colormap;
 } else {
  frag.rgb *= 0.75;
 }
 
 // alpha
 if (hasBumpAlpha > 0) {
  frag.a = normalmap.a;
 } else {
  frag.a = colormap.a;
 }
 
 // wreck map
 if (hasWreck > 0) {
  if (showDiffuse > 0) {
   frag.rgb *= wreckmap.rgb;
  }
 }
 
 // normal
 vec3 n;
 if (hasBump > 0) {
  // normal map
  n = normalize(normalmap.rgb * 2.0 - 1.0);
 } else {
  // vertex normal
  n = vertNorm;
 }
 
 // specular
 if (hasAlpha > 0) {
  if (hasBump > 0) {
   if (hasBumpAlpha > 0) {
    spec *= colormap.a;
   } else {
    spec *= normalmap.a;
   }
  }
 } else {
  spec *= colormap.a;
 }
 if (hasWreck > 0) {
  spec *= wreckmap.rgb;
 }
 
 // glass
 if (hasAlpha > 0) {
  if (hasEnvMap > 0) {
   vec3 v = normalize( eyepos - fragpos );
   frag.a += max(0.0, 1.0-dot(vertNorm,v));
  }
 }
 
 // lighting
 float NdotL = dot(n,normalize(-sunvec));
 if (showLighting > 0) {
  frag.rgb *= sunambient.rgb + sundiffuse.rgb * max(NdotL,0.0);
 }
 
 // envmap
 if (hasEnvMap > 0) {
  vec3 v = normalize( eyepos - fragpos );
  vec3 ref = reflect(v, vertNorm);
  vec4 env = textureCube(envmap, ref);
  frag.rgb = mix(frag.rgb, env.rgb, colormap.a);
 }
 
 // specular
 if (showLighting > 0) {
  if (NdotL > 0.0) {
   
   //spec = vec3(1.0, 0.0, 0.0);
   
   // half vector
   vec3 hv = normalize( -sunvec + eyesurfvec );
   
   // compute specular amount
   float NdotHV = max(dot(n,hv),0.0);
   float specPow = pow(NdotHV,100.0);
   spec *= specPow;
   
   // apply specular
   frag.rgb += sunspecular * spec;
   
   // glass specular
   if (hasEnvMap > 0) {
    if (hasAlpha > 0) {
     frag.a += specPow;
    }
   }
   
  }
 }
 
 // output
 gl_FragColor = frag;
 
 //gl_FragColor = vec4(vec3(normalize( eyepos - fragpos )),1.0);
 //gl_FragColor = vec4(vertNorm*0.5+0.5,1.0);
 //gl_FragColor.rgb = vec3(debug * 0.5 + 0.5);
}
