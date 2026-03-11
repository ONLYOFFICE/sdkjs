/*
 * (c) Copyright Ascensio System SIA 2010-2024
 *
 * This program is a free software product. You can redistribute it and/or
 * modify it under the terms of the GNU Affero General Public License (AGPL)
 * version 3 as published by the Free Software Foundation. In accordance with
 * Section 7(a) of the GNU AGPL its Section 15 shall be amended to the effect
 * that Ascensio System SIA expressly excludes the warranty of non-infringement
 * of any third-party rights.
 *
 * This program is distributed WITHOUT ANY WARRANTY; without even the implied
 * warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR  PURPOSE. For
 * details, see the GNU AGPL at: http://www.gnu.org/licenses/agpl-3.0.html
 *
 * You can contact Ascensio System SIA at 20A-6 Ernesta Birznieka-Upish
 * street, Riga, Latvia, EU, LV-1050.
 *
 * The  interactive user interfaces in modified source and object code versions
 * of the Program must display Appropriate Legal Notices, as required under
 * Section 5 of the GNU AGPL version 3.
 *
 * Pursuant to Section 7(b) of the License you must retain the original Product
 * logo when distributing the program. Pursuant to Section 7(e) we decline to
 * grant you any rights under trademark law for use of our trademarks.
 *
 * All the Product's GUI elements, including illustrations and icon sets, as
 * well as technical writing content are licensed under the terms of the
 * Creative Commons Attribution-ShareAlike 4.0 International. See the License
 * terms at http://creativecommons.org/licenses/by-sa/4.0/legalcode
 *
 */

"use strict";

// ============================================================
// CTransitionGL — WebGL-based slide transition renderer
// ============================================================

var c_oAscSlideTransitionTypes = Asc.c_oAscSlideTransitionTypes;
var c_oAscSlideTransitionParams = Asc.c_oAscSlideTransitionParams;

function CTransitionGL(transitionAnimation)
{
    this.TransitionAnimation = transitionAnimation;
    this.gl = null;
    this.glCanvas = null;
    this.programs = {};
    this.textures = { slide1: null, slide2: null };
    this.buffers = {};
    this.isInitialized = false;
    this.currentTransition = null;
}

// ============================================================
// Init / Dispose
// ============================================================

CTransitionGL.prototype.Init = function(w, h)
{
    if (this.glCanvas && this.gl && !this.gl.isContextLost() &&
        this.glCanvas.width === w && this.glCanvas.height === h)
    {
        this.isInitialized = true;
        return true;
    }

    // New context — old GL objects become invalid
    this.programs = {};
    this.buffers = {};
    this.textures = { slide1: null, slide2: null };

    this.glCanvas = document.createElement('canvas');
    this.glCanvas.width = w;
    this.glCanvas.height = h;

    let opts = {
        alpha: true,
        premultipliedAlpha: false,
        preserveDrawingBuffer: true,
        antialias: true,
        depth: true
    };

    this.gl = this.glCanvas.getContext('webgl', opts) || this.glCanvas.getContext('experimental-webgl', opts);
    if (!this.gl)
        return false;

    let gl = this.gl;
    gl.enable(gl.DEPTH_TEST);
    gl.depthFunc(gl.LEQUAL);
    gl.enable(gl.BLEND);
    gl.blendFunc(gl.SRC_ALPHA, gl.ONE_MINUS_SRC_ALPHA);

    // Handle context loss
    let oThis = this;
    this.glCanvas.addEventListener('webglcontextlost', function(e)
    {
        e.preventDefault();
        oThis.isInitialized = false;
        if (oThis.TransitionAnimation)
        {
            oThis.TransitionAnimation.IsWebGLAvailable = false;
        }
    }, false);

    this._initQuadBuffer();
    this.isInitialized = true;
    return true;
};

CTransitionGL.prototype.Dispose = function()
{
    if (!this.gl) return;
    let gl = this.gl;

    if (this.textures.slide1) { gl.deleteTexture(this.textures.slide1); this.textures.slide1 = null; }
    if (this.textures.slide2) { gl.deleteTexture(this.textures.slide2); this.textures.slide2 = null; }

    for (let name in this.programs)
    {
        if (this.programs.hasOwnProperty(name))
            gl.deleteProgram(this.programs[name].program);
    }
    this.programs = {};

    for (let name in this.buffers)
    {
        if (this.buffers.hasOwnProperty(name))
        {
            let buf = this.buffers[name];
            if (buf.vbo) gl.deleteBuffer(buf.vbo);
            if (buf.ibo) gl.deleteBuffer(buf.ibo);
        }
    }
    this.buffers = {};

    this.gl = null;
    this.glCanvas = null;
    this.isInitialized = false;
};

CTransitionGL.prototype._resetGLState = function()
{
    let gl = this.gl;
    if (!gl) return;

    // Restore masks
    gl.depthMask(true);
    gl.colorMask(true, true, true, true);

    // Restore Init() defaults: DEPTH_TEST + BLEND on, others off
    gl.enable(gl.DEPTH_TEST);
    gl.depthFunc(gl.LEQUAL);
    gl.enable(gl.BLEND);
    gl.blendFunc(gl.SRC_ALPHA, gl.ONE_MINUS_SRC_ALPHA);
    gl.disable(gl.CULL_FACE);
    gl.disable(gl.SCISSOR_TEST);
    gl.disable(gl.STENCIL_TEST);

    // Unbind buffers
    gl.bindBuffer(gl.ARRAY_BUFFER, null);
    gl.bindBuffer(gl.ELEMENT_ARRAY_BUFFER, null);

    // Disable all vertex attrib arrays
    let maxAttribs = gl.getParameter(gl.MAX_VERTEX_ATTRIBS);
    for (let i = 0; i < maxAttribs; i++)
        gl.disableVertexAttribArray(i);

    // Unbind textures and program
    gl.activeTexture(gl.TEXTURE1);
    gl.bindTexture(gl.TEXTURE_2D, null);
    gl.activeTexture(gl.TEXTURE0);
    gl.bindTexture(gl.TEXTURE_2D, null);
    gl.useProgram(null);
};

CTransitionGL.prototype.Cleanup = function()
{
    if (!this.gl) return;
    let gl = this.gl;

    this._resetGLState();

    if (this.textures.slide1) { gl.deleteTexture(this.textures.slide1); this.textures.slide1 = null; }
    if (this.textures.slide2) { gl.deleteTexture(this.textures.slide2); this.textures.slide2 = null; }
    this.currentTransition = null;
};

// ============================================================
// Texture management
// ============================================================

CTransitionGL.prototype.UploadSlideTextures = function(image1, image2)
{
    let gl = this.gl;
    if (!gl) return;

    if (this.textures.slide1) gl.deleteTexture(this.textures.slide1);
    if (this.textures.slide2) gl.deleteTexture(this.textures.slide2);

    this.textures.slide1 = image1 ? this._createTextureFromCanvas(image1) : null;
    this.textures.slide2 = image2 ? this._createTextureFromCanvas(image2) : null;
};

CTransitionGL.prototype._createTextureFromCanvas = function(canvas)
{
    let gl = this.gl;
    let tex = gl.createTexture();
    gl.bindTexture(gl.TEXTURE_2D, tex);
    gl.pixelStorei(gl.UNPACK_FLIP_Y_WEBGL, true);
    gl.texImage2D(gl.TEXTURE_2D, 0, gl.RGBA, gl.RGBA, gl.UNSIGNED_BYTE, canvas);
    gl.texParameteri(gl.TEXTURE_2D, gl.TEXTURE_WRAP_S, gl.CLAMP_TO_EDGE);
    gl.texParameteri(gl.TEXTURE_2D, gl.TEXTURE_WRAP_T, gl.CLAMP_TO_EDGE);
    gl.texParameteri(gl.TEXTURE_2D, gl.TEXTURE_MIN_FILTER, gl.LINEAR);
    gl.texParameteri(gl.TEXTURE_2D, gl.TEXTURE_MAG_FILTER, gl.LINEAR);
    return tex;
};

// ============================================================
// Shader compilation
// ============================================================

CTransitionGL.prototype.GetProgram = function(name, vertSrc, fragSrc)
{
    if (this.programs[name])
        return this.programs[name];

    let gl = this.gl;
    let vert = this._compileShader(gl.VERTEX_SHADER, vertSrc);
    let frag = this._compileShader(gl.FRAGMENT_SHADER, fragSrc);
    if (!vert || !frag) return null;

    let prog = gl.createProgram();
    gl.attachShader(prog, vert);
    gl.attachShader(prog, frag);
    gl.linkProgram(prog);

    if (!gl.getProgramParameter(prog, gl.LINK_STATUS))
    {
        console.error('Program link error: ' + gl.getProgramInfoLog(prog));
        gl.deleteProgram(prog);
        return null;
    }

    // Cache attribute and uniform locations
    let info = { program: prog, attrs: {}, uniforms: {} };

    let numAttrs = gl.getProgramParameter(prog, gl.ACTIVE_ATTRIBUTES);
    for (let i = 0; i < numAttrs; i++)
    {
        let attr = gl.getActiveAttrib(prog, i);
        info.attrs[attr.name] = gl.getAttribLocation(prog, attr.name);
    }

    let numUniforms = gl.getProgramParameter(prog, gl.ACTIVE_UNIFORMS);
    for (let i = 0; i < numUniforms; i++)
    {
        let uni = gl.getActiveUniform(prog, i);
        info.uniforms[uni.name] = gl.getUniformLocation(prog, uni.name);
    }

    this.programs[name] = info;
    return info;
};

CTransitionGL.prototype._compileShader = function(type, source)
{
    let gl = this.gl;
    let shader = gl.createShader(type);
    gl.shaderSource(shader, source);
    gl.compileShader(shader);
    if (!gl.getShaderParameter(shader, gl.COMPILE_STATUS))
    {
        console.error('Shader compile error: ' + gl.getShaderInfoLog(shader));
        gl.deleteShader(shader);
        return null;
    }
    return shader;
};

// ============================================================
// Geometry: full-screen quad
// ============================================================

CTransitionGL.prototype._initQuadBuffer = function()
{
    let gl = this.gl;
    // 2 triangles forming a full-screen quad
    // positions (x,y) + texcoords (u,v)
    let data = new Float32Array([
        // x,    y,   u,   v
        -1.0, -1.0,  0.0, 0.0,
         1.0, -1.0,  1.0, 0.0,
        -1.0,  1.0,  0.0, 1.0,
         1.0,  1.0,  1.0, 1.0
    ]);
    let vbo = gl.createBuffer();
    gl.bindBuffer(gl.ARRAY_BUFFER, vbo);
    gl.bufferData(gl.ARRAY_BUFFER, data, gl.STATIC_DRAW);
    this.buffers['quad'] = { vbo: vbo, vertCount: 4 };
};

CTransitionGL.prototype._bindQuad = function(progInfo)
{
    let gl = this.gl;
    let buf = this.buffers['quad'];
    gl.bindBuffer(gl.ARRAY_BUFFER, buf.vbo);

    let aPos = progInfo.attrs['aPosition'];
    let aTex = progInfo.attrs['aTexCoord'];

    if (aPos !== undefined && aPos !== -1)
    {
        gl.enableVertexAttribArray(aPos);
        gl.vertexAttribPointer(aPos, 2, gl.FLOAT, false, 16, 0);
    }
    if (aTex !== undefined && aTex !== -1)
    {
        gl.enableVertexAttribArray(aTex);
        gl.vertexAttribPointer(aTex, 2, gl.FLOAT, false, 16, 8);
    }
};

// ============================================================
// 4x4 Matrix utilities (column-major for WebGL)
// ============================================================

let _Mat4 = {
    identity: function()
    {
        return new Float32Array([
            1, 0, 0, 0,
            0, 1, 0, 0,
            0, 0, 1, 0,
            0, 0, 0, 1
        ]);
    },

    perspective: function(fovRad, aspect, near, far)
    {
        let f = 1.0 / Math.tan(fovRad / 2);
        let nf = 1.0 / (near - far);
        return new Float32Array([
            f / aspect, 0, 0, 0,
            0, f, 0, 0,
            0, 0, (far + near) * nf, -1,
            0, 0, 2 * far * near * nf, 0
        ]);
    },

    translate: function(m, tx, ty, tz)
    {
        let out = new Float32Array(m);
        out[12] += m[0] * tx + m[4] * ty + m[8]  * tz;
        out[13] += m[1] * tx + m[5] * ty + m[9]  * tz;
        out[14] += m[2] * tx + m[6] * ty + m[10] * tz;
        out[15] += m[3] * tx + m[7] * ty + m[11] * tz;
        return out;
    },

    rotateY: function(m, rad)
    {
        let s = Math.sin(rad), c = Math.cos(rad);
        let out = new Float32Array(m);
        let a00 = m[0], a01 = m[1], a02 = m[2], a03 = m[3];
        let a20 = m[8], a21 = m[9], a22 = m[10], a23 = m[11];
        out[0]  = a00 * c + a20 * s;
        out[1]  = a01 * c + a21 * s;
        out[2]  = a02 * c + a22 * s;
        out[3]  = a03 * c + a23 * s;
        out[8]  = a20 * c - a00 * s;
        out[9]  = a21 * c - a01 * s;
        out[10] = a22 * c - a02 * s;
        out[11] = a23 * c - a03 * s;
        return out;
    },

    rotateX: function(m, rad)
    {
        let s = Math.sin(rad), c = Math.cos(rad);
        let out = new Float32Array(m);
        let a10 = m[4], a11 = m[5], a12 = m[6], a13 = m[7];
        let a20 = m[8], a21 = m[9], a22 = m[10], a23 = m[11];
        out[4]  = a10 * c + a20 * s;
        out[5]  = a11 * c + a21 * s;
        out[6]  = a12 * c + a22 * s;
        out[7]  = a13 * c + a23 * s;
        out[8]  = a20 * c - a10 * s;
        out[9]  = a21 * c - a11 * s;
        out[10] = a22 * c - a12 * s;
        out[11] = a23 * c - a13 * s;
        return out;
    },

    rotateZ: function(m, rad)
    {
        let s = Math.sin(rad), c = Math.cos(rad);
        let out = new Float32Array(m);
        let a00 = m[0], a01 = m[1], a02 = m[2], a03 = m[3];
        let a10 = m[4], a11 = m[5], a12 = m[6], a13 = m[7];
        out[0]  = a00 * c + a10 * s;
        out[1]  = a01 * c + a11 * s;
        out[2]  = a02 * c + a12 * s;
        out[3]  = a03 * c + a13 * s;
        out[4]  = a10 * c - a00 * s;
        out[5]  = a11 * c - a01 * s;
        out[6]  = a12 * c - a02 * s;
        out[7]  = a13 * c - a03 * s;
        return out;
    },

    multiply: function(a, b)
    {
        let out = new Float32Array(16);
        for (let i = 0; i < 4; i++)
        {
            for (let j = 0; j < 4; j++)
            {
                out[i * 4 + j] = 0;
                for (let k = 0; k < 4; k++)
                    out[i * 4 + j] += b[k * 4 + j] * a[i * 4 + k];
            }
        }
        return out;
    }
};

// ============================================================
// Transition: PrepareTransition / Render dispatch
// ============================================================

CTransitionGL.prototype.PrepareTransition = function(type, param)
{
    this.currentTransition = { type: type, param: param };

    switch (type)
    {
        case c_oAscSlideTransitionTypes.Flip:
            this._prepareFlip();
            break;
        case c_oAscSlideTransitionTypes.Doors:
        case c_oAscSlideTransitionTypes.Window:
            this._prepareDoors();
            break;
        case c_oAscSlideTransitionTypes.Switch:
            this._prepareSwitch();
            break;
        case c_oAscSlideTransitionTypes.Gallery:
            this._prepareGallery();
            break;
        case c_oAscSlideTransitionTypes.Ferris:
            this._prepareFerris();
            break;
        case c_oAscSlideTransitionTypes.Prism:
            this._preparePrism();
            break;
        case c_oAscSlideTransitionTypes.Ripple:
            this._prepareRipple();
            break;
        case c_oAscSlideTransitionTypes.Vortex:
            this._prepareVortex();
            break;
        case c_oAscSlideTransitionTypes.Honeycomb:
            this._prepareHoneycomb();
            break;
        case c_oAscSlideTransitionTypes.Blinds:
            this._prepareBlinds();
            break;
        case c_oAscSlideTransitionTypes.Checker:
            this._prepareChecker();
            break;
        case c_oAscSlideTransitionTypes.Circle:
            this._prepareCircle();
            break;
        case c_oAscSlideTransitionTypes.Diamond:
            this._prepareDiamond();
            break;
        case c_oAscSlideTransitionTypes.Plus:
            this._preparePlus();
            break;
        case c_oAscSlideTransitionTypes.RandomBar:
            this._prepareRandomBar();
            break;
        case c_oAscSlideTransitionTypes.Dissolve:
            this._prepareDissolve();
            break;
        default:
            this._prepareCrossfade();
            break;
    }
};

CTransitionGL.prototype.Render = function(type, param, progress)
{
    if (!this.gl || !this.isInitialized) return;

    window['_dbgGL'] = this;
    window['_dbgType'] = type;
    window['_dbgParam'] = param;

    let gl = this.gl;

    this._resetGLState();

    gl.viewport(0, 0, this.glCanvas.width, this.glCanvas.height);
    gl.clearColor(0, 0, 0, 1);
    gl.clear(gl.COLOR_BUFFER_BIT | gl.DEPTH_BUFFER_BIT);

    switch (type)
    {
        case c_oAscSlideTransitionTypes.Flip:
            this._renderFlip(progress, param);
            break;
        case c_oAscSlideTransitionTypes.Doors:
            this._renderDoors(progress, param, false);
            break;
        case c_oAscSlideTransitionTypes.Window:
            this._renderDoors(progress, param, true);
            break;
        case c_oAscSlideTransitionTypes.Switch:
            this._renderSwitch(progress, param);
            break;
        case c_oAscSlideTransitionTypes.Gallery:
            this._renderGallery(progress, param);
            break;
        case c_oAscSlideTransitionTypes.Ferris:
            this._renderFerris(progress, param);
            break;
        case c_oAscSlideTransitionTypes.Prism:
            this._renderPrism(progress, param);
            break;
        case c_oAscSlideTransitionTypes.Ripple:
            this._renderRipple(progress, param);
            break;
        case c_oAscSlideTransitionTypes.Vortex:
            this._renderVortex(progress, param);
            break;
        case c_oAscSlideTransitionTypes.Honeycomb:
            this._renderHoneycomb(progress);
            break;
        case c_oAscSlideTransitionTypes.Blinds:
            this._renderBlinds(progress, param);
            break;
        case c_oAscSlideTransitionTypes.Checker:
            this._renderChecker(progress, param);
            break;
        case c_oAscSlideTransitionTypes.Circle:
            this._renderCircle(progress);
            break;
        case c_oAscSlideTransitionTypes.Diamond:
            this._renderDiamond(progress);
            break;
        case c_oAscSlideTransitionTypes.Plus:
            this._renderPlus(progress);
            break;
        case c_oAscSlideTransitionTypes.RandomBar:
            this._renderRandomBar(progress, param);
            break;
        case c_oAscSlideTransitionTypes.Dissolve:
            this._renderDissolve(progress);
            break;
        default:
            this._renderCrossfade(progress);
            break;
    }
};

// ============================================================
// Common shaders
// ============================================================

let _VERT_3D = [
    'attribute vec3 aPosition;',
    'attribute vec2 aTexCoord;',
    'uniform mat4 uProjection;',
    'uniform mat4 uModelView;',
    'varying vec2 vTexCoord;',
    'void main() {',
    '    gl_Position = uProjection * uModelView * vec4(aPosition, 1.0);',
    '    vTexCoord = aTexCoord;',
    '}'
].join('\n');

let _FRAG_TEXTURED = [
    'precision mediump float;',
    'uniform sampler2D uTexture;',
    'uniform float uAlpha;',
    'varying vec2 vTexCoord;',
    'void main() {',
    '    vec4 color = texture2D(uTexture, vTexCoord);',
    '    gl_FragColor = vec4(color.rgb, color.a * uAlpha);',
    '}'
].join('\n');

let _FRAG_GALLERY_REFL = [
    'precision mediump float;',
    'uniform sampler2D uTexture;',
    'uniform float uAlpha;',
    'varying vec2 vTexCoord;',
    'void main() {',
    '    vec4 color = texture2D(uTexture, vTexCoord);',
    '    float grad = 1.0 - vTexCoord.y;',
    '    gl_FragColor = vec4(color.rgb, color.a * uAlpha * grad * grad);',
    '}'
].join('\n');

let _VERT_QUAD = [
    'attribute vec2 aPosition;',
    'attribute vec2 aTexCoord;',
    'varying vec2 vTexCoord;',
    'void main() {',
    '    vTexCoord = aTexCoord;',
    '    gl_Position = vec4(aPosition, 0.0, 1.0);',
    '}'
].join('\n');

let _FRAG_CROSSFADE = [
    'precision mediump float;',
    'uniform sampler2D uTexture1;',
    'uniform sampler2D uTexture2;',
    'uniform float uProgress;',
    'varying vec2 vTexCoord;',
    'void main() {',
    '    vec4 c1 = texture2D(uTexture1, vTexCoord);',
    '    vec4 c2 = texture2D(uTexture2, vTexCoord);',
    '    gl_FragColor = mix(c1, c2, uProgress);',
    '}'
].join('\n');

// ============================================================
// Transition: Crossfade (fallback for unimplemented WebGL transitions)
// ============================================================

CTransitionGL.prototype._prepareCrossfade = function()
{
    this.GetProgram('crossfade', _VERT_QUAD, _FRAG_CROSSFADE);
};

CTransitionGL.prototype._renderCrossfade = function(progress)
{
    let gl = this.gl;
    let prog = this.programs['crossfade'];
    if (!prog) return;

    gl.useProgram(prog.program);
    gl.disable(gl.DEPTH_TEST);

    gl.activeTexture(gl.TEXTURE0);
    gl.bindTexture(gl.TEXTURE_2D, this.textures.slide1);
    gl.uniform1i(prog.uniforms['uTexture1'], 0);

    gl.activeTexture(gl.TEXTURE1);
    gl.bindTexture(gl.TEXTURE_2D, this.textures.slide2);
    gl.uniform1i(prog.uniforms['uTexture2'], 1);

    gl.uniform1f(prog.uniforms['uProgress'], progress);

    this._bindQuad(prog);
    gl.drawArrays(gl.TRIANGLE_STRIP, 0, 4);
};

// ============================================================
// Transition: Flip — Y-axis card flip
// ============================================================

CTransitionGL.prototype._prepareFlip = function()
{
    this.GetProgram('flip3d', _VERT_3D, _FRAG_TEXTURED);
    this._initQuadBuffer3D();
};

CTransitionGL.prototype._initQuadBuffer3D = function()
{
    if (this.buffers['quad3d']) return;

    let gl = this.gl;
    let aspect = this.glCanvas.width / this.glCanvas.height;
    let hw = aspect, hh = 1.0;

    // 3D quad: position(x,y,z) + texcoord(u,v) = 5 floats per vertex
    let data = new Float32Array([
        -hw, -hh, 0,  0, 0,
         hw, -hh, 0,  1, 0,
        -hw,  hh, 0,  0, 1,
         hw,  hh, 0,  1, 1
    ]);
    let vbo = gl.createBuffer();
    gl.bindBuffer(gl.ARRAY_BUFFER, vbo);
    gl.bufferData(gl.ARRAY_BUFFER, data, gl.STATIC_DRAW);
    this.buffers['quad3d'] = { vbo: vbo, vertCount: 4, aspect: aspect };
};

CTransitionGL.prototype._bindQuad3D = function(progInfo)
{
    let gl = this.gl;
    let buf = this.buffers['quad3d'];
    gl.bindBuffer(gl.ARRAY_BUFFER, buf.vbo);

    let stride = 20; // 5 floats * 4 bytes
    let aPos = progInfo.attrs['aPosition'];
    let aTex = progInfo.attrs['aTexCoord'];

    if (aPos !== undefined && aPos !== -1)
    {
        gl.enableVertexAttribArray(aPos);
        gl.vertexAttribPointer(aPos, 3, gl.FLOAT, false, stride, 0);
    }
    if (aTex !== undefined && aTex !== -1)
    {
        gl.enableVertexAttribArray(aTex);
        gl.vertexAttribPointer(aTex, 2, gl.FLOAT, false, stride, 12);
    }
};

CTransitionGL.prototype._renderFlip = function(progress, param)
{
    let gl = this.gl;
    let prog = this.programs['flip3d'];
    if (!prog) return;

    gl.useProgram(prog.program);
    gl.enable(gl.DEPTH_TEST);

    let aspect = this.glCanvas.width / this.glCanvas.height;
    let fov = Math.PI / 4;
    let dist = 1.0 / Math.tan(fov / 2);
    let projection = _Mat4.perspective(fov, aspect, 0.1, 100.0);

    let isLeft = (param === c_oAscSlideTransitionParams.Flip_Left);
    let dir = isLeft ? -1 : 1;
    let angle = dir * progress * Math.PI; // 0 to ±180°

    gl.uniformMatrix4fv(prog.uniforms['uProjection'], false, projection);
    gl.uniform1f(prog.uniforms['uAlpha'], 1.0);

    if (progress <= 0.5)
    {
        // Old slide: front face rotating away (0 → ±90°)
        let mv = _Mat4.identity();
        mv = _Mat4.translate(mv, 0, 0, -dist);
        mv = _Mat4.rotateY(mv, angle);

        gl.uniformMatrix4fv(prog.uniforms['uModelView'], false, mv);
        gl.activeTexture(gl.TEXTURE0);
        gl.bindTexture(gl.TEXTURE_2D, this.textures.slide1);
        gl.uniform1i(prog.uniforms['uTexture'], 0);
        this._bindQuad3D(prog);
        gl.drawArrays(gl.TRIANGLE_STRIP, 0, 4);
    }
    else
    {
        // New slide: back face rotating into view (∓90° → 0)
        let backAngle = angle - dir * Math.PI;
        let mv = _Mat4.identity();
        mv = _Mat4.translate(mv, 0, 0, -dist);
        mv = _Mat4.rotateY(mv, backAngle);

        gl.uniformMatrix4fv(prog.uniforms['uModelView'], false, mv);
        gl.activeTexture(gl.TEXTURE0);
        gl.bindTexture(gl.TEXTURE_2D, this.textures.slide2);
        gl.uniform1i(prog.uniforms['uTexture'], 0);
        this._bindQuad3D(prog);
        gl.drawArrays(gl.TRIANGLE_STRIP, 0, 4);
    }
};

// ============================================================
// Transition: Doors — two halves swing open
// ============================================================

CTransitionGL.prototype._prepareDoors = function()
{
    this.GetProgram('flip3d', _VERT_3D, _FRAG_TEXTURED);
    this._initHalfQuadBuffers();
};

CTransitionGL.prototype._initHalfQuadBuffers = function()
{
    if (this.buffers['halfLeft']) return;

    let gl = this.gl;
    let aspect = this.glCanvas.width / this.glCanvas.height;
    let hw = aspect, hh = 1.0;

    // Left half (hinge at left edge)
    let leftData = new Float32Array([
        -hw, -hh, 0,  0.0, 0.0,
         0,  -hh, 0,  0.5, 0.0,
        -hw,  hh, 0,  0.0, 1.0,
         0,   hh, 0,  0.5, 1.0
    ]);
    let leftVbo = gl.createBuffer();
    gl.bindBuffer(gl.ARRAY_BUFFER, leftVbo);
    gl.bufferData(gl.ARRAY_BUFFER, leftData, gl.STATIC_DRAW);
    this.buffers['halfLeft'] = { vbo: leftVbo, vertCount: 4 };

    // Right half (hinge at right edge)
    let rightData = new Float32Array([
         0,  -hh, 0,  0.5, 0.0,
         hw, -hh, 0,  1.0, 0.0,
         0,   hh, 0,  0.5, 1.0,
         hw,  hh, 0,  1.0, 1.0
    ]);
    let rightVbo = gl.createBuffer();
    gl.bindBuffer(gl.ARRAY_BUFFER, rightVbo);
    gl.bufferData(gl.ARRAY_BUFFER, rightData, gl.STATIC_DRAW);
    this.buffers['halfRight'] = { vbo: rightVbo, vertCount: 4 };

    // Top half (hinge at top edge)
    let topData = new Float32Array([
        -hw,  0,  0,  0.0, 0.5,
         hw,  0,  0,  1.0, 0.5,
        -hw,  hh, 0,  0.0, 1.0,
         hw,  hh, 0,  1.0, 1.0
    ]);
    let topVbo = gl.createBuffer();
    gl.bindBuffer(gl.ARRAY_BUFFER, topVbo);
    gl.bufferData(gl.ARRAY_BUFFER, topData, gl.STATIC_DRAW);
    this.buffers['halfTop'] = { vbo: topVbo, vertCount: 4 };

    // Bottom half (hinge at bottom edge)
    let bottomData = new Float32Array([
        -hw, -hh, 0,  0.0, 0.0,
         hw, -hh, 0,  1.0, 0.0,
        -hw,  0,  0,  0.0, 0.5,
         hw,  0,  0,  1.0, 0.5
    ]);
    let bottomVbo = gl.createBuffer();
    gl.bindBuffer(gl.ARRAY_BUFFER, bottomVbo);
    gl.bufferData(gl.ARRAY_BUFFER, bottomData, gl.STATIC_DRAW);
    this.buffers['halfBottom'] = { vbo: bottomVbo, vertCount: 4 };
};

CTransitionGL.prototype._bindBuffer3D = function(bufName, progInfo)
{
    let gl = this.gl;
    let buf = this.buffers[bufName];
    if (!buf) return;
    gl.bindBuffer(gl.ARRAY_BUFFER, buf.vbo);

    let stride = 20;
    let aPos = progInfo.attrs['aPosition'];
    let aTex = progInfo.attrs['aTexCoord'];
    if (aPos !== undefined && aPos !== -1)
    {
        gl.enableVertexAttribArray(aPos);
        gl.vertexAttribPointer(aPos, 3, gl.FLOAT, false, stride, 0);
    }
    if (aTex !== undefined && aTex !== -1)
    {
        gl.enableVertexAttribArray(aTex);
        gl.vertexAttribPointer(aTex, 2, gl.FLOAT, false, stride, 12);
    }
};

CTransitionGL.prototype._renderDoors = function(progress, param, isWindow)
{
    let gl = this.gl;
    let prog = this.programs['flip3d'];
    if (!prog) return;

    gl.useProgram(prog.program);
    gl.enable(gl.DEPTH_TEST);

    let aspect = this.glCanvas.width / this.glCanvas.height;
    let fov = Math.PI / 4;
    let dist = 1.0 / Math.tan(fov / 2);
    let projection = _Mat4.perspective(fov, aspect, 0.1, 100.0);

    let isHorz = (param === c_oAscSlideTransitionParams.Doors_Horizontal ||
                  param === c_oAscSlideTransitionParams.Window_Horizontal);
    let sign = isWindow ? 1 : -1;

    // Doors take ~85% of the time; new slide zooms for the full duration
    let doorEnd = 0.9;
    let doorProgress = Math.min(progress / doorEnd, 1.0);
    let angle = doorProgress * Math.PI / 2;

    // Draw new slide behind (full quad) — starts at ~half size and zooms in
    this._initQuadBuffer3D();
    {
        let mv = _Mat4.identity();
        mv = _Mat4.translate(mv, 0, 0, -dist * (2 - progress));

        gl.uniformMatrix4fv(prog.uniforms['uProjection'], false, projection);
        gl.uniformMatrix4fv(prog.uniforms['uModelView'], false, mv);
        gl.uniform1f(prog.uniforms['uAlpha'], 1.0);

        gl.activeTexture(gl.TEXTURE0);
        gl.bindTexture(gl.TEXTURE_2D, this.textures.slide2);
        gl.uniform1i(prog.uniforms['uTexture'], 0);

        this._bindQuad3D(prog);
        gl.drawArrays(gl.TRIANGLE_STRIP, 0, 4);
    }

    // Draw old slide halves (only while doors are still open)
    if (doorProgress < 1.0)
    {
        let halfAlpha = 1.0 - doorProgress;

        // Keep framebuffer alpha at 1.0 so the 2D layer below doesn't show through
        gl.blendFuncSeparate(gl.SRC_ALPHA, gl.ONE_MINUS_SRC_ALPHA, gl.ONE, gl.ONE_MINUS_SRC_ALPHA);

        // Left/top half
        {
            let hw = aspect;
            let mv = _Mat4.identity();
            mv = _Mat4.translate(mv, 0, 0, -dist);
            if (isHorz)
            {
                mv = _Mat4.translate(mv, 0, 1.0, 0);
                mv = _Mat4.rotateX(mv, sign * angle);
                mv = _Mat4.translate(mv, 0, -1.0, 0);
            }
            else
            {
                mv = _Mat4.translate(mv, -hw, 0, 0);
                mv = _Mat4.rotateY(mv, -sign * angle);
                mv = _Mat4.translate(mv, hw, 0, 0);
            }

            gl.uniformMatrix4fv(prog.uniforms['uProjection'], false, projection);
            gl.uniformMatrix4fv(prog.uniforms['uModelView'], false, mv);
            gl.uniform1f(prog.uniforms['uAlpha'], halfAlpha);

            gl.activeTexture(gl.TEXTURE0);
            gl.bindTexture(gl.TEXTURE_2D, this.textures.slide1);
            gl.uniform1i(prog.uniforms['uTexture'], 0);

            this._bindBuffer3D(isHorz ? 'halfTop' : 'halfLeft', prog);
            gl.drawArrays(gl.TRIANGLE_STRIP, 0, 4);
        }

        // Right/bottom half
        {
            let hw = aspect;
            let mv = _Mat4.identity();
            mv = _Mat4.translate(mv, 0, 0, -dist);
            if (isHorz)
            {
                mv = _Mat4.translate(mv, 0, -1.0, 0);
                mv = _Mat4.rotateX(mv, -sign * angle);
                mv = _Mat4.translate(mv, 0, 1.0, 0);
            }
            else
            {
                mv = _Mat4.translate(mv, hw, 0, 0);
                mv = _Mat4.rotateY(mv, sign * angle);
                mv = _Mat4.translate(mv, -hw, 0, 0);
            }

            gl.uniformMatrix4fv(prog.uniforms['uProjection'], false, projection);
            gl.uniformMatrix4fv(prog.uniforms['uModelView'], false, mv);
            gl.uniform1f(prog.uniforms['uAlpha'], halfAlpha);

            gl.activeTexture(gl.TEXTURE0);
            gl.bindTexture(gl.TEXTURE_2D, this.textures.slide1);
            gl.uniform1i(prog.uniforms['uTexture'], 0);

            this._bindBuffer3D(isHorz ? 'halfBottom' : 'halfRight', prog);
            gl.drawArrays(gl.TRIANGLE_STRIP, 0, 4);
        }

        // Restore standard blend func
        gl.blendFunc(gl.SRC_ALPHA, gl.ONE_MINUS_SRC_ALPHA);
    }
};

// ============================================================
// Transition: Switch — two-card carousel swap
// ============================================================

CTransitionGL.prototype._prepareSwitch = function()
{
    this.GetProgram('flip3d', _VERT_3D, _FRAG_TEXTURED);
    this._initQuadBuffer3D();
};

CTransitionGL.prototype._renderSwitch = function(progress, param)
{
    let gl = this.gl;
    let prog = this.programs['flip3d'];
    if (!prog) return;

    gl.useProgram(prog.program);
    gl.enable(gl.DEPTH_TEST);

    let aspect = this.glCanvas.width / this.glCanvas.height;
    let fov = Math.PI / 4;
    let dist = 1.0 / Math.tan(fov / 2);
    let projection = _Mat4.perspective(fov, aspect, 0.1, 100.0);

    let isLeft = (param === c_oAscSlideTransitionParams.Switch_Left);
    let dir = isLeft ? -1 : 1;

    // 90-degree carousel: slides form an L-shape, rotating together
    let angle = progress * Math.PI / 2; // 0 to 90 degrees
    let R = aspect; // carousel radius = slide half-width

    // Old slide rotating away
    {
        let a = dir * angle;
        let x = Math.sin(a) * R;
        let z = -dist + (Math.cos(a) - 1) * R;
        let mv = _Mat4.identity();
        mv = _Mat4.translate(mv, x, 0, z);
        mv = _Mat4.rotateY(mv, a);

        gl.uniformMatrix4fv(prog.uniforms['uProjection'], false, projection);
        gl.uniformMatrix4fv(prog.uniforms['uModelView'], false, mv);
        gl.uniform1f(prog.uniforms['uAlpha'], 1.0);
        gl.activeTexture(gl.TEXTURE0);
        gl.bindTexture(gl.TEXTURE_2D, this.textures.slide1);
        gl.uniform1i(prog.uniforms['uTexture'], 0);
        this._bindQuad3D(prog);
        gl.drawArrays(gl.TRIANGLE_STRIP, 0, 4);
    }

    // New slide rotating in (starts at -90°, ends at 0°)
    {
        let a = dir * (angle - Math.PI / 2);
        let x = Math.sin(a) * R;
        let z = -dist + (Math.cos(a) - 1) * R;
        let mv = _Mat4.identity();
        mv = _Mat4.translate(mv, x, 0, z);
        mv = _Mat4.rotateY(mv, a);

        gl.uniformMatrix4fv(prog.uniforms['uProjection'], false, projection);
        gl.uniformMatrix4fv(prog.uniforms['uModelView'], false, mv);
        gl.uniform1f(prog.uniforms['uAlpha'], 1.0);
        gl.activeTexture(gl.TEXTURE0);
        gl.bindTexture(gl.TEXTURE_2D, this.textures.slide2);
        gl.uniform1i(prog.uniforms['uTexture'], 0);
        this._bindQuad3D(prog);
        gl.drawArrays(gl.TRIANGLE_STRIP, 0, 4);
    }
};

// ============================================================
// Transition: Gallery — perspective sliding
// ============================================================

CTransitionGL.prototype._prepareGallery = function()
{
    this.GetProgram('flip3d', _VERT_3D, _FRAG_TEXTURED);
    this.GetProgram('galleryRefl', _VERT_3D, _FRAG_GALLERY_REFL);
    this._initQuadBuffer3D();
};

CTransitionGL.prototype._renderGallerySlide = function(prog, reflProg, projection, texture, t, dir, targetBottom)
{
    let gl = this.gl;
    let aspect = this.glCanvas.width / this.glCanvas.height;
    let dist = 1.0 / Math.tan(Math.PI / 8);
    let hh = 1.0;

    let maxZ = 3.0;
    let tiltMax = Math.PI / 6;
    let xFactor = 2.2;

    let extraZ = t * maxZ;
    let scale = dist / (dist + extraZ);
    // Anchor bottom edge at targetBottom NDC line
    let yShift = targetBottom / scale + 1.0;
    let slideX = dir * t * aspect * xFactor;
    let tiltAngle = dir * t * tiltMax;
    let slideZ = -dist - extraZ;

    let mv = _Mat4.identity();
    mv = _Mat4.translate(mv, slideX, yShift, slideZ);
    mv = _Mat4.rotateY(mv, tiltAngle);

    // Main slide
    gl.useProgram(prog.program);
    gl.uniformMatrix4fv(prog.uniforms['uProjection'], false, projection);
    gl.uniformMatrix4fv(prog.uniforms['uModelView'], false, mv);
    gl.uniform1f(prog.uniforms['uAlpha'], 1.0);
    gl.activeTexture(gl.TEXTURE0);
    gl.bindTexture(gl.TEXTURE_2D, texture);
    gl.uniform1i(prog.uniforms['uTexture'], 0);
    this._bindQuad3D(prog);
    gl.drawArrays(gl.TRIANGLE_STRIP, 0, 4);

    // Reflection below
    if (reflProg && t > 0.01)
    {
        gl.useProgram(reflProg.program);
        gl.blendFuncSeparate(gl.SRC_ALPHA, gl.ONE_MINUS_SRC_ALPHA, gl.ONE, gl.ONE_MINUS_SRC_ALPHA);

        let mvRefl = _Mat4.translate(mv, 0, -2.0 * hh, 0);
        mvRefl[4] = -mvRefl[4];
        mvRefl[5] = -mvRefl[5];
        mvRefl[6] = -mvRefl[6];
        mvRefl[7] = -mvRefl[7];

        gl.uniformMatrix4fv(reflProg.uniforms['uProjection'], false, projection);
        gl.uniformMatrix4fv(reflProg.uniforms['uModelView'], false, mvRefl);
        gl.uniform1f(reflProg.uniforms['uAlpha'], 0.3);
        gl.activeTexture(gl.TEXTURE0);
        gl.bindTexture(gl.TEXTURE_2D, texture);
        gl.uniform1i(reflProg.uniforms['uTexture'], 0);
        this._bindQuad3D(reflProg);
        gl.drawArrays(gl.TRIANGLE_STRIP, 0, 4);

        gl.blendFunc(gl.SRC_ALPHA, gl.ONE_MINUS_SRC_ALPHA);
    }
};

CTransitionGL.prototype._renderGallery = function(progress, param)
{
    let gl = this.gl;
    let prog = this.programs['flip3d'];
    let reflProg = this.programs['galleryRefl'];
    if (!prog) return;

    gl.enable(gl.DEPTH_TEST);

    let fov = Math.PI / 4;
    let aspect = this.glCanvas.width / this.glCanvas.height;
    let projection = _Mat4.perspective(fov, aspect, 0.1, 100.0);

    let isLeft = (param === c_oAscSlideTransitionParams.Gallery_Left);
    let dir = isLeft ? -1 : 1;

    // Common bottom line: rises quickly from -1.0, stays constant, drops back
    // Both slides' bottom edges travel along this same line
    let bottomRise = Math.min(progress * 4, (1 - progress) * 4, 1.0) * 0.35;
    let targetBottom = -1.0 + bottomRise;

    // Old slide: t goes 0 → 1 (moves away)
    this._renderGallerySlide(prog, reflProg, projection, this.textures.slide1, progress, dir, targetBottom);

    // New slide: t goes 1 → 0 (arrives), opposite direction
    this._renderGallerySlide(prog, reflProg, projection, this.textures.slide2, 1.0 - progress, -dir, targetBottom);
};

// ============================================================
// Geometry: subdivided mesh for deformation transitions
// ============================================================

CTransitionGL.prototype._initMeshBuffer = function(segsX, segsY, name)
{
    if (this.buffers[name]) return;

    let gl = this.gl;
    let aspect = this.glCanvas.width / this.glCanvas.height;
    let hw = aspect, hh = 1.0;

    let vertCount = (segsX + 1) * (segsY + 1);
    let idxCount = segsX * segsY * 6;
    // position(x,y,z) + texcoord(u,v) = 5 floats
    let verts = new Float32Array(vertCount * 5);
    let indices = new Uint16Array(idxCount);

    let vi = 0;
    for (let iy = 0; iy <= segsY; iy++)
    {
        let v = iy / segsY;
        let y = -hh + v * 2 * hh;
        for (let ix = 0; ix <= segsX; ix++)
        {
            let u = ix / segsX;
            let x = -hw + u * 2 * hw;
            verts[vi++] = x;
            verts[vi++] = y;
            verts[vi++] = 0;
            verts[vi++] = u;
            verts[vi++] = v;
        }
    }

    let ii = 0;
    for (let iy = 0; iy < segsY; iy++)
    {
        for (let ix = 0; ix < segsX; ix++)
        {
            let a = iy * (segsX + 1) + ix;
            let b = a + 1;
            let c = a + (segsX + 1);
            let d = c + 1;
            indices[ii++] = a; indices[ii++] = b; indices[ii++] = c;
            indices[ii++] = b; indices[ii++] = d; indices[ii++] = c;
        }
    }

    let vbo = gl.createBuffer();
    gl.bindBuffer(gl.ARRAY_BUFFER, vbo);
    gl.bufferData(gl.ARRAY_BUFFER, verts, gl.STATIC_DRAW);

    let ibo = gl.createBuffer();
    gl.bindBuffer(gl.ELEMENT_ARRAY_BUFFER, ibo);
    gl.bufferData(gl.ELEMENT_ARRAY_BUFFER, indices, gl.STATIC_DRAW);

    this.buffers[name] = { vbo: vbo, ibo: ibo, idxCount: idxCount, vertCount: vertCount };
};

CTransitionGL.prototype._bindMesh = function(name, progInfo)
{
    let gl = this.gl;
    let buf = this.buffers[name];
    if (!buf) return;

    gl.bindBuffer(gl.ARRAY_BUFFER, buf.vbo);
    gl.bindBuffer(gl.ELEMENT_ARRAY_BUFFER, buf.ibo);

    let stride = 20;
    let aPos = progInfo.attrs['aPosition'];
    let aTex = progInfo.attrs['aTexCoord'];
    if (aPos !== undefined && aPos !== -1)
    {
        gl.enableVertexAttribArray(aPos);
        gl.vertexAttribPointer(aPos, 3, gl.FLOAT, false, stride, 0);
    }
    if (aTex !== undefined && aTex !== -1)
    {
        gl.enableVertexAttribArray(aTex);
        gl.vertexAttribPointer(aTex, 2, gl.FLOAT, false, stride, 12);
    }
};

// ============================================================
// Transition: Ripple — wave displacement
// ============================================================

CTransitionGL.prototype._prepareRipple = function()
{
    this.GetProgram('ripple', _VERT_QUAD, _FRAG_RIPPLE);
};

CTransitionGL.prototype._renderRipple = function(progress, param)
{
    let gl = this.gl;
    let prog = this.programs['ripple'];
    if (!prog) return;

    gl.useProgram(prog.program);
    gl.disable(gl.DEPTH_TEST);

    // Origin in UV space (0,0)=bottom-left, (1,1)=top-right
    let originX = 0.5, originY = 0.5;
    switch (param)
    {
        case c_oAscSlideTransitionParams.Ripple_LeftUp:    originX = 1.0; originY = 0.0; break;
        case c_oAscSlideTransitionParams.Ripple_RightUp:   originX = 0.0; originY = 0.0; break;
        case c_oAscSlideTransitionParams.Ripple_LeftDown:  originX = 1.0; originY = 1.0; break;
        case c_oAscSlideTransitionParams.Ripple_RightDown: originX = 0.0; originY = 1.0; break;
        case c_oAscSlideTransitionParams.Ripple_Center:    originX = 0.5; originY = 0.5; break;
    }

    gl.activeTexture(gl.TEXTURE0);
    gl.bindTexture(gl.TEXTURE_2D, this.textures.slide1);
    gl.uniform1i(prog.uniforms['uTexture1'], 0);

    gl.activeTexture(gl.TEXTURE1);
    gl.bindTexture(gl.TEXTURE_2D, this.textures.slide2);
    gl.uniform1i(prog.uniforms['uTexture2'], 1);

    gl.uniform1f(prog.uniforms['uProgress'], progress);
    gl.uniform1f(prog.uniforms['uAspect'], this.glCanvas.width / this.glCanvas.height);
    gl.uniform2f(prog.uniforms['uOrigin'], originX, originY);

    this._bindQuad(prog);
    gl.drawArrays(gl.TRIANGLE_STRIP, 0, 4);
};

CTransitionGL.prototype._prepareFerris = function()
{
    this.GetProgram('flip3d', _VERT_3D, _FRAG_TEXTURED);
    this._initQuadBuffer3D();
};

CTransitionGL.prototype._renderFerris = function(progress, param)
{
    let gl = this.gl;
    let prog = this.programs['flip3d'];
    if (!prog) return;

    gl.useProgram(prog.program);
    gl.enable(gl.DEPTH_TEST);

    let aspect = this.glCanvas.width / this.glCanvas.height;
    let fov = Math.PI / 4;
    let dist = 1.0 / Math.tan(fov / 2);
    let projection = _Mat4.perspective(fov, aspect, 0.1, 100.0);

    let isLeft = (param === c_oAscSlideTransitionParams.Ferris_Left);
    let dir = isLeft ? 1 : -1;
    let arcRadius = 2.0;
    let arcAngle = progress * Math.PI / 2;

    // Old slide rotates away on arc
    {
        let x = dir * Math.sin(arcAngle) * arcRadius;
        let y = -(1.0 - Math.cos(arcAngle)) * arcRadius;
        let tilt = dir * arcAngle;

        let mv = _Mat4.identity();
        mv = _Mat4.translate(mv, x, y, -dist);
        mv = _Mat4.rotateY(mv, 0);
        // Tilt the slide as it goes around the wheel
        mv[0] = Math.cos(tilt);  mv[4] = -Math.sin(tilt);
        mv[1] = Math.sin(tilt);  mv[5] = Math.cos(tilt);

        let mvFull = _Mat4.identity();
        mvFull = _Mat4.translate(mvFull, x, y, -dist);
        // Apply Z-rotation for ferris wheel tilt
        let cosT = Math.cos(tilt), sinT = Math.sin(tilt);
        let rotZ = _Mat4.identity();
        rotZ[0] = cosT; rotZ[4] = -sinT;
        rotZ[1] = sinT; rotZ[5] = cosT;
        mvFull = _Mat4.multiply(mvFull, rotZ);

        gl.uniformMatrix4fv(prog.uniforms['uProjection'], false, projection);
        gl.uniformMatrix4fv(prog.uniforms['uModelView'], false, mvFull);
        gl.uniform1f(prog.uniforms['uAlpha'], 1.0 - progress);
        gl.activeTexture(gl.TEXTURE0);
        gl.bindTexture(gl.TEXTURE_2D, this.textures.slide1);
        gl.uniform1i(prog.uniforms['uTexture'], 0);
        this._bindQuad3D(prog);
        gl.drawArrays(gl.TRIANGLE_STRIP, 0, 4);
    }

    // New slide rotates in from opposite side
    {
        let revAngle = (1.0 - progress) * Math.PI / 2;
        let x = -dir * Math.sin(revAngle) * arcRadius;
        let y = -(1.0 - Math.cos(revAngle)) * arcRadius;
        let tilt = -dir * revAngle;

        let mvFull = _Mat4.identity();
        mvFull = _Mat4.translate(mvFull, x, y, -dist);
        let cosT = Math.cos(tilt), sinT = Math.sin(tilt);
        let rotZ = _Mat4.identity();
        rotZ[0] = cosT; rotZ[4] = -sinT;
        rotZ[1] = sinT; rotZ[5] = cosT;
        mvFull = _Mat4.multiply(mvFull, rotZ);

        gl.uniformMatrix4fv(prog.uniforms['uProjection'], false, projection);
        gl.uniformMatrix4fv(prog.uniforms['uModelView'], false, mvFull);
        gl.uniform1f(prog.uniforms['uAlpha'], progress);
        gl.activeTexture(gl.TEXTURE0);
        gl.bindTexture(gl.TEXTURE_2D, this.textures.slide2);
        gl.uniform1i(prog.uniforms['uTexture'], 0);
        this._bindQuad3D(prog);
        gl.drawArrays(gl.TRIANGLE_STRIP, 0, 4);
    }
};

// ============================================================
// Transition: Prism — rectangular prism (cube) rotation, 90°
// ============================================================

CTransitionGL.prototype._preparePrism = function()
{
    this.GetProgram('flip3d', _VERT_3D, _FRAG_TEXTURED);
    this._initQuadBuffer3D();
};

CTransitionGL.prototype._renderPrism = function(progress, param)
{
    let gl = this.gl;
    let prog = this.programs['flip3d'];
    if (!prog) return;

    gl.useProgram(prog.program);
    gl.enable(gl.DEPTH_TEST);

    let aspect = this.glCanvas.width / this.glCanvas.height;
    let fov = Math.PI / 4;
    let dist = 1.0 / Math.tan(fov / 2);
    let projection = _Mat4.perspective(fov, aspect, 0.1, 100.0);

    // Decode direction from param
    let baseParam = c_oAscSlideTransitionParams.Prism_Left;
    let offset = param - baseParam;
    let dirIdx = offset % 4;  // 0=left, 1=right, 2=up, 3=down
    let isInverted = (offset >= 4 && offset < 8) || (offset >= 12);

    // Prism rotates 90 degrees (rectangular prism / cube face-to-face)
    let angle = progress * (Math.PI / 2);
    // Distance from prism center to face = half-width of the face
    let prismR = (dirIdx <= 1) ? aspect : 1.0;

    let isVertical = (dirIdx <= 1); // left/right rotate around Y
    let dirSign = (dirIdx === 0 || dirIdx === 2) ? 1 : -1;

    // isInverted=0 (default): faces on OUTSIDE (convex cube, axis behind faces)
    // isInverted=1: faces on INSIDE (concave, axis in front of faces)
    let pivotZ = isInverted ? prismR : -prismR;

    // Old slide face
    {
        let mv = _Mat4.identity();
        mv = _Mat4.translate(mv, 0, 0, -dist);
        if (isVertical)
        {
            mv = _Mat4.translate(mv, 0, 0, pivotZ);
            mv = _Mat4.rotateY(mv, dirSign * angle);
            mv = _Mat4.translate(mv, 0, 0, -pivotZ);
        }
        else
        {
            mv = _Mat4.translate(mv, 0, 0, pivotZ);
            mv = _Mat4.rotateX(mv, -dirSign * angle);
            mv = _Mat4.translate(mv, 0, 0, -pivotZ);
        }

        gl.uniformMatrix4fv(prog.uniforms['uProjection'], false, projection);
        gl.uniformMatrix4fv(prog.uniforms['uModelView'], false, mv);
        gl.uniform1f(prog.uniforms['uAlpha'], 1.0);
        gl.activeTexture(gl.TEXTURE0);
        gl.bindTexture(gl.TEXTURE_2D, this.textures.slide1);
        gl.uniform1i(prog.uniforms['uTexture'], 0);
        this._bindQuad3D(prog);
        gl.drawArrays(gl.TRIANGLE_STRIP, 0, 4);
    }

    // New slide face (offset by 90 degrees on the prism)
    {
        let faceAngle = Math.PI / 2; // 90 degrees between adjacent cube faces
        let mv = _Mat4.identity();
        mv = _Mat4.translate(mv, 0, 0, -dist);
        if (isVertical)
        {
            mv = _Mat4.translate(mv, 0, 0, pivotZ);
            mv = _Mat4.rotateY(mv, dirSign * (angle - faceAngle));
            mv = _Mat4.translate(mv, 0, 0, -pivotZ);
        }
        else
        {
            mv = _Mat4.translate(mv, 0, 0, pivotZ);
            mv = _Mat4.rotateX(mv, -dirSign * (angle - faceAngle));
            mv = _Mat4.translate(mv, 0, 0, -pivotZ);
        }

        gl.uniformMatrix4fv(prog.uniforms['uProjection'], false, projection);
        gl.uniformMatrix4fv(prog.uniforms['uModelView'], false, mv);
        gl.uniform1f(prog.uniforms['uAlpha'], 1.0);
        gl.activeTexture(gl.TEXTURE0);
        gl.bindTexture(gl.TEXTURE_2D, this.textures.slide2);
        gl.uniform1i(prog.uniforms['uTexture'], 0);
        this._bindQuad3D(prog);
        gl.drawArrays(gl.TRIANGLE_STRIP, 0, 4);
    }
};

let _FRAG_TILE_SCATTER = [
    'precision mediump float;',
    'uniform sampler2D uTexture;',
    'varying vec2 vTexCoord;',
    'varying float vAlpha;',
    'void main() {',
    '    vec4 color = texture2D(uTexture, vTexCoord);',
    '    gl_FragColor = vec4(color.rgb, color.a * vAlpha);',
    '}'
].join('\n');

// ============================================================
// Tile mesh generation for scatter transitions
// ============================================================

CTransitionGL.prototype._initTileMeshBuffer = function(tilesX, tilesY, name)
{
    if (this.buffers[name]) return;

    let gl = this.gl;
    let aspect = this.glCanvas.width / this.glCanvas.height;
    let hw = aspect, hh = 1.0;

    let tileCount = tilesX * tilesY;
    // Each tile is 2 triangles = 6 vertices
    // Each vertex: position(3) + texcoord(2) + tileOffset(3) + tilePhase(1) + tileCenter(2) = 11 floats
    let vertCount = tileCount * 6;
    let data = new Float32Array(vertCount * 11);

    let seed = 12345;
    let rand = function() { seed = (seed * 16807 + 0) % 2147483647; return (seed & 0x7fffffff) / 2147483647; };

    let vi = 0;
    for (let ty = 0; ty < tilesY; ty++)
    {
        for (let tx = 0; tx < tilesX; tx++)
        {
            let u0 = tx / tilesX, u1 = (tx + 1) / tilesX;
            let v0 = ty / tilesY, v1 = (ty + 1) / tilesY;
            let x0 = -hw + u0 * 2 * hw, x1 = -hw + u1 * 2 * hw;
            let y0 = -hh + v0 * 2 * hh, y1 = -hh + v1 * 2 * hh;

            let centerX = (x0 + x1) * 0.5;
            let centerY = (y0 + y1) * 0.5;

            let offX = (rand() - 0.5) * 2.0;
            let offY = (rand() - 0.5) * 2.0;
            let offZ = rand() * 0.5 + 0.5;
            let phase = rand();

            let corners = [
                [x0, y0, 0, u0, v0],
                [x1, y0, 0, u1, v0],
                [x0, y1, 0, u0, v1],
                [x1, y0, 0, u1, v0],
                [x1, y1, 0, u1, v1],
                [x0, y1, 0, u0, v1]
            ];
            for (let c = 0; c < 6; c++)
            {
                data[vi++] = corners[c][0];
                data[vi++] = corners[c][1];
                data[vi++] = corners[c][2];
                data[vi++] = corners[c][3];
                data[vi++] = corners[c][4];
                data[vi++] = offX;
                data[vi++] = offY;
                data[vi++] = offZ;
                data[vi++] = phase;
                data[vi++] = centerX;
                data[vi++] = centerY;
            }
        }
    }

    let vbo = gl.createBuffer();
    gl.bindBuffer(gl.ARRAY_BUFFER, vbo);
    gl.bufferData(gl.ARRAY_BUFFER, data, gl.STATIC_DRAW);
    this.buffers[name] = { vbo: vbo, vertCount: vertCount, stride: 44 };
};

CTransitionGL.prototype._bindTileMesh = function(name, progInfo)
{
    let gl = this.gl;
    let buf = this.buffers[name];
    if (!buf) return;

    gl.bindBuffer(gl.ARRAY_BUFFER, buf.vbo);
    let stride = buf.stride;

    let aPos = progInfo.attrs['aPosition'];
    let aTex = progInfo.attrs['aTexCoord'];
    let aOff = progInfo.attrs['aTileOffset'];
    let aPh  = progInfo.attrs['aTilePhase'];
    let aCtr = progInfo.attrs['aTileCenter'];

    if (aPos !== undefined && aPos !== -1)
    {
        gl.enableVertexAttribArray(aPos);
        gl.vertexAttribPointer(aPos, 3, gl.FLOAT, false, stride, 0);
    }
    if (aTex !== undefined && aTex !== -1)
    {
        gl.enableVertexAttribArray(aTex);
        gl.vertexAttribPointer(aTex, 2, gl.FLOAT, false, stride, 12);
    }
    if (aOff !== undefined && aOff !== -1)
    {
        gl.enableVertexAttribArray(aOff);
        gl.vertexAttribPointer(aOff, 3, gl.FLOAT, false, stride, 20);
    }
    if (aPh !== undefined && aPh !== -1)
    {
        gl.enableVertexAttribArray(aPh);
        gl.vertexAttribPointer(aPh, 1, gl.FLOAT, false, stride, 32);
    }
    if (aCtr !== undefined && aCtr !== -1)
    {
        gl.enableVertexAttribArray(aCtr);
        gl.vertexAttribPointer(aCtr, 2, gl.FLOAT, false, stride, 36);
    }
};

// ============================================================
// Transition: Vortex — fragment shader vortex effect
// ============================================================

let _VERT_VORTEX_SCATTER = [
    'attribute vec3 aPosition;',
    'attribute vec2 aTexCoord;',
    'attribute vec3 aTileOffset;',
    'attribute float aTilePhase;',
    'attribute vec2 aTileCenter;',
    'uniform mat4 uProjection;',
    'uniform mat4 uModelView;',
    'uniform float uProgress;',
    'uniform float uDirection;',
    'uniform float uAspect;',
    'uniform float uReverse;',
    'varying vec2 vTexCoord;',
    'varying float vAlpha;',
    '',
    'vec3 rotAx(vec3 v, vec3 ax, float a) {',
    '    float c = cos(a); float s = sin(a);',
    '    vec3 n = normalize(ax);',
    '    return v * c + cross(n, v) * s + n * dot(n, v) * (1.0 - c);',
    '}',
    '',
    'void main() {',
    '    vec3 pos = aPosition;',
    '    vec3 local = pos - vec3(aTileCenter, 0.0);',
    '',
    '    // Direction-based wave coordinate',
    '    float coord;',
    '    float maxR;',
    '    if (uDirection < 0.5) {',
    '        coord = aTileCenter.x; maxR = uAspect;',
    '    } else if (uDirection < 1.5) {',
    '        coord = -aTileCenter.x; maxR = uAspect;',
    '    } else if (uDirection < 2.5) {',
    '        coord = -aTileCenter.y; maxR = 1.0;',
    '    } else {',
    '        coord = aTileCenter.y; maxR = 1.0;',
    '    }',
    '',
    '    float normalized = (coord + maxR) / (2.0 * maxR);',
    '    if (uReverse > 0.5) normalized = 1.0 - normalized;',
    '',
    '    float delay = (1.0 - normalized) * 0.5 + aTilePhase * 0.08;',
    '    float t = smoothstep(delay, delay + 0.4, uProgress);',
    '    float tAnim = (uReverse > 0.5) ? 1.0 - t : t;',
    '',
    '    vAlpha = 1.0 - smoothstep(0.85, 1.0, tAnim);',
    '',
    '    // Y-axis rotation: tile centers orbit into depth',
    '    float rotAngle = tAnim * 1.5708;',
    '    float rSign;',
    '    if (uDirection < 0.5) rSign = 1.0;',
    '    else if (uDirection < 1.5) rSign = -1.0;',
    '    else if (uDirection < 2.5) rSign = 1.0;',
    '    else rSign = -1.0;',
    '',
    '    float cr = cos(rotAngle * rSign);',
    '    float sr = sin(rotAngle * rSign);',
    '',
    '    // Rotate both tile center and local geometry around same axis',
    '    vec3 rotPos;',
    '    if (uDirection < 1.5) {',
    '        local = vec3(local.x * cr, local.y, -local.x * sr);',
    '        rotPos = vec3(aTileCenter.x * cr, aTileCenter.y, -aTileCenter.x * sr);',
    '    } else {',
    '        local = vec3(local.x, local.y * cr, -local.y * sr);',
    '        rotPos = vec3(aTileCenter.x, aTileCenter.y * cr, -aTileCenter.y * sr);',
    '    }',
    '',
    '    pos = rotPos + local;',
    '',
    '    // Cloud scatter + directional drift',
    '    pos.x += aTileOffset.x * tAnim * 1.2;',
    '    pos.y += aTileOffset.y * tAnim * 1.0;',
    '    pos.z += (aTileOffset.z - 0.75) * tAnim * 5.0;',
    '',
    '    // Drift: "Left" = particles move right (away from wave), etc.',
    '    if (uDirection < 0.5) pos.x += tAnim * 0.6;',
    '    else if (uDirection < 1.5) pos.x -= tAnim * 0.6;',
    '    else if (uDirection < 2.5) pos.y -= tAnim * 0.6;',
    '    else pos.y += tAnim * 0.6;',
    '',
    '    vTexCoord = aTexCoord;',
    '    gl_Position = uProjection * uModelView * vec4(pos, 1.0);',
    '}'
].join('\n');

CTransitionGL.prototype._prepareVortex = function()
{
    this.GetProgram('vortexScatter', _VERT_VORTEX_SCATTER, _FRAG_TILE_SCATTER);
    this._initTileMeshBuffer(200, 112, 'tilesVortex');
};

CTransitionGL.prototype._renderVortex = function(progress, param)
{
    let gl = this.gl;
    let prog = this.programs['vortexScatter'];
    if (!prog) return;

    let tileBuf = this.buffers['tilesVortex'];
    if (!tileBuf) return;

    let aspect = this.glCanvas.width / this.glCanvas.height;
    let fov = Math.PI / 4;
    let dist = 1.0 / Math.tan(fov / 2);
    let projection = _Mat4.perspective(fov, aspect, 0.1, 100.0);
    let mv = _Mat4.translate(_Mat4.identity(), 0, 0, -dist);

    let dir = 0;
    if (param === c_oAscSlideTransitionParams.Vortex_Right) dir = 1;
    else if (param === c_oAscSlideTransitionParams.Vortex_Up) dir = 2;
    else if (param === c_oAscSlideTransitionParams.Vortex_Down) dir = 3;

    gl.clearColor(0.0, 0.0, 0.0, 1.0);
    gl.clear(gl.COLOR_BUFFER_BIT | gl.DEPTH_BUFFER_BIT);
    gl.enable(gl.DEPTH_TEST);
    gl.enable(gl.BLEND);
    gl.blendFunc(gl.SRC_ALPHA, gl.ONE_MINUS_SRC_ALPHA);

    gl.useProgram(prog.program);
    gl.uniformMatrix4fv(prog.uniforms['uProjection'], false, projection);
    gl.uniformMatrix4fv(prog.uniforms['uModelView'], false, mv);
    gl.uniform1f(prog.uniforms['uDirection'], dir);
    gl.uniform1f(prog.uniforms['uAspect'], aspect);

    // Split timeline: old scatters 0→0.5, new assembles 0.5→1.0
    let oldProgress = Math.min(1.0, progress * 2.0);
    let newProgress = Math.max(0.0, (progress - 0.5) * 2.0);

    // Pass 1: new slide tiles assembling (draw first = behind)
    gl.depthMask(false);
    gl.uniform1f(prog.uniforms['uProgress'], newProgress);
    gl.uniform1f(prog.uniforms['uReverse'], 1.0);
    gl.activeTexture(gl.TEXTURE0);
    gl.bindTexture(gl.TEXTURE_2D, this.textures.slide2);
    gl.uniform1i(prog.uniforms['uTexture'], 0);
    this._bindTileMesh('tilesVortex', prog);
    gl.drawArrays(gl.TRIANGLES, 0, tileBuf.vertCount);

    // Pass 2: old slide tiles scattering (draw second = in front)
    gl.depthMask(true);
    gl.uniform1f(prog.uniforms['uProgress'], oldProgress);
    gl.uniform1f(prog.uniforms['uReverse'], 0.0);
    gl.activeTexture(gl.TEXTURE0);
    gl.bindTexture(gl.TEXTURE_2D, this.textures.slide1);
    gl.uniform1i(prog.uniforms['uTexture'], 0);
    gl.drawArrays(gl.TRIANGLES, 0, tileBuf.vertCount);

    gl.disable(gl.BLEND);
    gl.disable(gl.DEPTH_TEST);
};

// ============================================================
// Transition: Ripple — fullscreen 2D UV-distortion ripple
// ============================================================

let _FRAG_RIPPLE = [
    'precision mediump float;',
    'uniform sampler2D uTexture1;',
    'uniform sampler2D uTexture2;',
    'uniform float uProgress;',
    'uniform float uAspect;',
    'uniform vec2 uOrigin;',
    'varying vec2 vTexCoord;',
    'void main() {',
    '    vec2 uv = vTexCoord;',
    '    vec2 diff = uv - uOrigin;',
    '    diff.x *= uAspect;',
    '    float dist = length(diff);',
    '    float maxR = length(vec2(uAspect, 1.0));',
    '    float wavefrontR = uProgress * (maxR + 0.15);',
    '    float d = dist - wavefrontR;',
    '    float envelope = exp(-d * d * 6.0);',
    '    float wave = sin(d * 25.0);',
    '    vec2 dir = dist > 0.001 ? normalize(diff) : vec2(1.0, 0.0);',
    '    dir.x /= uAspect;',
    '    vec2 rippleUV = uv + dir * wave * envelope * 0.008;',
    '    vec4 c1 = texture2D(uTexture1, rippleUV);',
    '    vec4 c2 = texture2D(uTexture2, rippleUV);',
    '    float behind = smoothstep(0.03, -0.03, d);',
    '    float shade = 1.0 - 0.08 * (1.0 - wave) * 0.5 * envelope;',
    '    vec4 result = mix(c1, c2, behind);',
    '    gl_FragColor = vec4(result.rgb * shade, result.a);',
    '}'
].join('\n');

// ============================================================
// Transition: Circle — expanding circle with gradient edge
// ============================================================

let _FRAG_CIRCLE = [
    'precision mediump float;',
    'uniform sampler2D uTexture1;',
    'uniform sampler2D uTexture2;',
    'uniform float uProgress;',
    'uniform float uAspect;',
    'varying vec2 vTexCoord;',
    'void main() {',
    '    vec4 c1 = texture2D(uTexture1, vTexCoord);',
    '    vec4 c2 = texture2D(uTexture2, vTexCoord);',
    '    vec2 uv = vTexCoord - 0.5;',
    '    uv.x *= uAspect;',
    '    float maxR = length(vec2(0.5 * uAspect, 0.5));',
    '    float edgeW = maxR * 0.08;',
    '    float r = uProgress * (maxR + edgeW);',
    '    float d = length(uv);',
    '    float alpha = smoothstep(r, r - edgeW, d);',
    '    gl_FragColor = mix(c1, c2, alpha);',
    '}'
].join('\n');

CTransitionGL.prototype._prepareCircle = function()
{
    this.GetProgram('circle', _VERT_QUAD, _FRAG_CIRCLE);
};

CTransitionGL.prototype._renderCircle = function(progress)
{
    let gl = this.gl;
    let prog = this.programs['circle'];
    if (!prog) return;

    gl.useProgram(prog.program);
    gl.disable(gl.DEPTH_TEST);

    gl.activeTexture(gl.TEXTURE0);
    gl.bindTexture(gl.TEXTURE_2D, this.textures.slide1);
    gl.uniform1i(prog.uniforms['uTexture1'], 0);

    gl.activeTexture(gl.TEXTURE1);
    gl.bindTexture(gl.TEXTURE_2D, this.textures.slide2);
    gl.uniform1i(prog.uniforms['uTexture2'], 1);

    gl.uniform1f(prog.uniforms['uProgress'], progress);
    gl.uniform1f(prog.uniforms['uAspect'], this.glCanvas.width / this.glCanvas.height);

    this._bindQuad(prog);
    gl.drawArrays(gl.TRIANGLE_STRIP, 0, 4);
};

// ============================================================
// Transition: Diamond — expanding diamond with gradient edge
// ============================================================

let _FRAG_DIAMOND = [
    'precision mediump float;',
    'uniform sampler2D uTexture1;',
    'uniform sampler2D uTexture2;',
    'uniform float uProgress;',
    'varying vec2 vTexCoord;',
    'void main() {',
    '    vec4 c1 = texture2D(uTexture1, vTexCoord);',
    '    vec4 c2 = texture2D(uTexture2, vTexCoord);',
    '    float maxR = 1.0;',
    '    float edgeW = maxR * 0.08;',
    '    float r = uProgress * (maxR + edgeW);',
    '    float d = abs(vTexCoord.x - 0.5) + abs(vTexCoord.y - 0.5);',
    '    float alpha = smoothstep(r, r - edgeW, d);',
    '    gl_FragColor = mix(c1, c2, alpha);',
    '}'
].join('\n');

CTransitionGL.prototype._prepareDiamond = function()
{
    this.GetProgram('diamond', _VERT_QUAD, _FRAG_DIAMOND);
};

CTransitionGL.prototype._renderDiamond = function(progress)
{
    let gl = this.gl;
    let prog = this.programs['diamond'];
    if (!prog) return;

    gl.useProgram(prog.program);
    gl.disable(gl.DEPTH_TEST);

    gl.activeTexture(gl.TEXTURE0);
    gl.bindTexture(gl.TEXTURE_2D, this.textures.slide1);
    gl.uniform1i(prog.uniforms['uTexture1'], 0);

    gl.activeTexture(gl.TEXTURE1);
    gl.bindTexture(gl.TEXTURE_2D, this.textures.slide2);
    gl.uniform1i(prog.uniforms['uTexture2'], 1);

    gl.uniform1f(prog.uniforms['uProgress'], progress);

    this._bindQuad(prog);
    gl.drawArrays(gl.TRIANGLE_STRIP, 0, 4);
};

// ============================================================
// Transition: Plus — cross shape expanding from center
// ============================================================

let _FRAG_PLUS = [
    'precision mediump float;',
    'uniform sampler2D uTexture1;',
    'uniform sampler2D uTexture2;',
    'uniform float uProgress;',
    'varying vec2 vTexCoord;',
    'void main() {',
    '    vec4 c1 = texture2D(uTexture1, vTexCoord);',
    '    vec4 c2 = texture2D(uTexture2, vTexCoord);',
    '    float nx = abs(vTexCoord.x - 0.5) * 2.0;',
    '    float ny = abs(vTexCoord.y - 0.5) * 2.0;',
    '    float d = min(nx, ny);',
    '    float edgeFrac = 0.06;',
    '    float part = uProgress * (1.0 + edgeFrac);',
    '    float alpha = smoothstep(part - edgeFrac, part, d);',
    '    gl_FragColor = mix(c2, c1, alpha);',
    '}'
].join('\n');

CTransitionGL.prototype._preparePlus = function()
{
    this.GetProgram('plus', _VERT_QUAD, _FRAG_PLUS);
};

CTransitionGL.prototype._renderPlus = function(progress)
{
    let gl = this.gl;
    let prog = this.programs['plus'];
    if (!prog) return;

    gl.useProgram(prog.program);
    gl.disable(gl.DEPTH_TEST);

    gl.activeTexture(gl.TEXTURE0);
    gl.bindTexture(gl.TEXTURE_2D, this.textures.slide1);
    gl.uniform1i(prog.uniforms['uTexture1'], 0);

    gl.activeTexture(gl.TEXTURE1);
    gl.bindTexture(gl.TEXTURE_2D, this.textures.slide2);
    gl.uniform1i(prog.uniforms['uTexture2'], 1);

    gl.uniform1f(prog.uniforms['uProgress'], progress);

    this._bindQuad(prog);
    gl.drawArrays(gl.TRIANGLE_STRIP, 0, 4);
};

// ============================================================
// Transition: RandomBar — random strips with gradient edges
// ============================================================

let _FRAG_RANDOMBAR = [
    'precision mediump float;',
    'uniform sampler2D uTexture1;',
    'uniform sampler2D uTexture2;',
    'uniform float uProgress;',
    'uniform float uStripCount;',
    'uniform float uIsVertical;',
    'varying vec2 vTexCoord;',
    'float hash(float n) { return fract(sin(n * 127.1) * 43758.5453); }',
    'void main() {',
    '    vec4 c1 = texture2D(uTexture1, vTexCoord);',
    '    vec4 c2 = texture2D(uTexture2, vTexCoord);',
    '    float coord = uIsVertical > 0.5 ? vTexCoord.x : vTexCoord.y;',
    '    float si = floor(coord * uStripCount);',
    '    if (si >= uStripCount) si = uStripCount - 1.0;',
    '    if (si < 0.0) si = 0.0;',
    '    float threshold = hash(si);',
    '    if (uProgress < threshold) { gl_FragColor = c1; return; }',
    '    float localPos = fract(coord * uStripCount);',
    '    float edgeW = 0.2;',
    '    float alpha = 1.0;',
    '    if (si > 0.0 && uProgress < hash(si - 1.0))',
    '        alpha = min(alpha, smoothstep(0.0, edgeW, localPos));',
    '    if (si < uStripCount - 1.0 && uProgress < hash(si + 1.0))',
    '        alpha = min(alpha, smoothstep(0.0, edgeW, 1.0 - localPos));',
    '    gl_FragColor = mix(c1, c2, alpha);',
    '}'
].join('\n');

CTransitionGL.prototype._prepareRandomBar = function()
{
    this.GetProgram('randombar', _VERT_QUAD, _FRAG_RANDOMBAR);
};

CTransitionGL.prototype._renderRandomBar = function(progress, param)
{
    let gl = this.gl;
    let prog = this.programs['randombar'];
    if (!prog) return;

    gl.useProgram(prog.program);
    gl.disable(gl.DEPTH_TEST);

    gl.activeTexture(gl.TEXTURE0);
    gl.bindTexture(gl.TEXTURE_2D, this.textures.slide1);
    gl.uniform1i(prog.uniforms['uTexture1'], 0);

    gl.activeTexture(gl.TEXTURE1);
    gl.bindTexture(gl.TEXTURE_2D, this.textures.slide2);
    gl.uniform1i(prog.uniforms['uTexture2'], 1);

    gl.uniform1f(prog.uniforms['uProgress'], progress);
    let isVert = (param === c_oAscSlideTransitionParams.RandomBar_Vertical);
    gl.uniform1f(prog.uniforms['uStripCount'], isVert ? 195.0 : 150.0);
    gl.uniform1f(prog.uniforms['uIsVertical'], isVert ? 1.0 : 0.0);

    this._bindQuad(prog);
    gl.drawArrays(gl.TRIANGLE_STRIP, 0, 4);
};

// ============================================================
// Transition: Dissolve — random cell grid reveal
// ============================================================

let _FRAG_DISSOLVE = [
    'precision mediump float;',
    'uniform sampler2D uTexture1;',
    'uniform sampler2D uTexture2;',
    'uniform float uProgress;',
    'uniform float uCols;',
    'uniform float uRows;',
    'varying vec2 vTexCoord;',
    'float hash(vec2 p) { return fract(sin(dot(p, vec2(12.9898, 78.233))) * 43758.5453); }',
    'void main() {',
    '    vec4 c1 = texture2D(uTexture1, vTexCoord);',
    '    vec4 c2 = texture2D(uTexture2, vTexCoord);',
    '    vec2 cell = floor(vTexCoord * vec2(uCols, uRows));',
    '    float threshold = hash(cell);',
    '    gl_FragColor = uProgress >= threshold ? c2 : c1;',
    '}'
].join('\n');

CTransitionGL.prototype._prepareDissolve = function()
{
    this.GetProgram('dissolve', _VERT_QUAD, _FRAG_DISSOLVE);
};

CTransitionGL.prototype._renderDissolve = function(progress)
{
    let gl = this.gl;
    let prog = this.programs['dissolve'];
    if (!prog) return;

    gl.useProgram(prog.program);
    gl.disable(gl.DEPTH_TEST);

    gl.activeTexture(gl.TEXTURE0);
    gl.bindTexture(gl.TEXTURE_2D, this.textures.slide1);
    gl.uniform1i(prog.uniforms['uTexture1'], 0);

    gl.activeTexture(gl.TEXTURE1);
    gl.bindTexture(gl.TEXTURE_2D, this.textures.slide2);
    gl.uniform1i(prog.uniforms['uTexture2'], 1);

    gl.uniform1f(prog.uniforms['uProgress'], progress);
    gl.uniform1f(prog.uniforms['uCols'], 70.0);
    gl.uniform1f(prog.uniforms['uRows'], 55.0);

    this._bindQuad(prog);
    gl.drawArrays(gl.TRIANGLE_STRIP, 0, 4);
};

// ============================================================
// Transition: Honeycomb — hexagonal tiles with 3D flip reveal
// ============================================================

let _VERT_HONEYCOMB = [
    'attribute vec3 aPosition;',
    'attribute vec2 aTexCoord;',
    'attribute float aTilePhase;',
    'attribute vec2 aTileCenter;',
    'uniform mat4 uProjection;',
    'uniform mat4 uModelView;',
    'uniform float uProgress;',
    'varying vec2 vTexCoord;',
    'varying float vFlipProgress;',
    'varying float vShade;',
    'varying float vAlpha;',
    'void main() {',
    '    vec3 pos = aPosition;',
    '    float flipDuration = 0.45;',
    '    float flipStart = aTilePhase * (1.0 - flipDuration);',
    '    float fp = clamp((uProgress - flipStart) / flipDuration, 0.0, 1.0);',
    '    fp = fp * fp * (3.0 - 2.0 * fp);',
    '    vFlipProgress = fp;',
    '',
    '    vec3 local = pos - vec3(aTileCenter, 0.0);',
    '',
    '    float scale;',
    '    if (fp <= 0.5) {',
    '        scale = 1.0 - fp * 2.0;',
    '    } else {',
    '        scale = (fp - 0.5) * 2.0;',
    '    }',
    '    local.x *= scale;',
    '    local.y *= scale;',
    '    vShade = 1.0;',
    '    vAlpha = 1.0;',
    '',
    '    pos = vec3(aTileCenter, 0.0) + local;',
    '    vTexCoord = aTexCoord;',
    '    gl_Position = uProjection * uModelView * vec4(pos, 1.0);',
    '}'
].join('\n');

let _FRAG_HONEYCOMB = [
    'precision mediump float;',
    'uniform sampler2D uTexture1;',
    'uniform sampler2D uTexture2;',
    'varying vec2 vTexCoord;',
    'varying float vFlipProgress;',
    'varying float vShade;',
    'varying float vAlpha;',
    'void main() {',
    '    vec4 color;',
    '    if (vFlipProgress < 0.5) {',
    '        color = texture2D(uTexture1, vTexCoord);',
    '    } else {',
    '        color = texture2D(uTexture2, vTexCoord);',
    '    }',
    '    gl_FragColor = vec4(color.rgb * vShade, color.a * vAlpha);',
    '}'
].join('\n');

CTransitionGL.prototype._initHexMeshBuffer = function(name, hexRadius)
{
    if (this.buffers[name]) return;

    let gl = this.gl;
    let aspect = this.glCanvas.width / this.glCanvas.height;
    let hw = aspect, hh = 1.0;

    let r = hexRadius || 0.18;
    let colSpacing = r * Math.sqrt(3);
    let rowSpacing = r * 1.5;

    let cols = Math.ceil(hw * 2 / colSpacing) + 2;
    let rows = Math.ceil(hh * 2 / rowSpacing) + 2;

    let seed = 54321;
    let rand = function() { seed = (seed * 16807) % 2147483647; return (seed & 0x7fffffff) / 2147483647; };

    let hexes = [];
    for (let row = 0; row < rows; row++)
    {
        let cy = -hh - r + row * rowSpacing;
        for (let col = 0; col < cols; col++)
        {
            let cx = -hw - r + col * colSpacing + (row % 2) * colSpacing * 0.5;
            if (cx + r < -hw - 0.2 || cx - r > hw + 0.2) continue;
            if (cy + r < -hh - 0.2 || cy - r > hh + 0.2) continue;

            // Phase: diagonal wave from bottom-left to top-right
            let nx = (cx + hw) / (2 * hw);
            let ny = (cy + hh) / (2 * hh);
            let diag = (nx + ny) / 2;
            let phase = diag * 0.85 + rand() * 0.15;
            if (phase > 1.0) phase = 1.0;
            hexes.push({ cx: cx, cy: cy, phase: phase });
        }
    }

    // Each hex = 6 triangles x 3 vertices = 18 vertices
    // Vertex: position(3) + texcoord(2) + tilePhase(1) + tileCenter(2) = 8 floats
    let vertCount = hexes.length * 18;
    let data = new Float32Array(vertCount * 8);
    let vi = 0;

    let hexScale = 0.88; // visible gap between hexes

    for (let h = 0; h < hexes.length; h++)
    {
        let hex = hexes[h];
        let cx = hex.cx, cy = hex.cy;

        for (let tri = 0; tri < 6; tri++)
        {
            let a0 = (Math.PI / 6) + (tri * Math.PI / 3);
            let a1 = (Math.PI / 6) + ((tri + 1) * Math.PI / 3);

            let verts = [
                [cx, cy],
                [cx + r * hexScale * Math.cos(a0), cy + r * hexScale * Math.sin(a0)],
                [cx + r * hexScale * Math.cos(a1), cy + r * hexScale * Math.sin(a1)]
            ];

            for (let v = 0; v < 3; v++)
            {
                let px = verts[v][0];
                let py = verts[v][1];
                let u = (px + hw) / (2 * hw);
                let vv = (py + hh) / (2 * hh);

                data[vi++] = px;
                data[vi++] = py;
                data[vi++] = 0;
                data[vi++] = u;
                data[vi++] = vv;
                data[vi++] = hex.phase;
                data[vi++] = cx;
                data[vi++] = cy;
            }
        }
    }

    let vbo = gl.createBuffer();
    gl.bindBuffer(gl.ARRAY_BUFFER, vbo);
    gl.bufferData(gl.ARRAY_BUFFER, data, gl.STATIC_DRAW);
    this.buffers[name] = { vbo: vbo, vertCount: vertCount, stride: 32 };
};

CTransitionGL.prototype._bindHexMesh = function(name, progInfo)
{
    let gl = this.gl;
    let buf = this.buffers[name];
    if (!buf) return;

    gl.bindBuffer(gl.ARRAY_BUFFER, buf.vbo);
    let stride = buf.stride;

    let aPos = progInfo.attrs['aPosition'];
    let aTex = progInfo.attrs['aTexCoord'];
    let aPh  = progInfo.attrs['aTilePhase'];
    let aCtr = progInfo.attrs['aTileCenter'];

    if (aPos !== undefined && aPos !== -1)
    {
        gl.enableVertexAttribArray(aPos);
        gl.vertexAttribPointer(aPos, 3, gl.FLOAT, false, stride, 0);
    }
    if (aTex !== undefined && aTex !== -1)
    {
        gl.enableVertexAttribArray(aTex);
        gl.vertexAttribPointer(aTex, 2, gl.FLOAT, false, stride, 12);
    }
    if (aPh !== undefined && aPh !== -1)
    {
        gl.enableVertexAttribArray(aPh);
        gl.vertexAttribPointer(aPh, 1, gl.FLOAT, false, stride, 20);
    }
    if (aCtr !== undefined && aCtr !== -1)
    {
        gl.enableVertexAttribArray(aCtr);
        gl.vertexAttribPointer(aCtr, 2, gl.FLOAT, false, stride, 24);
    }
};

CTransitionGL.prototype._prepareHoneycomb = function()
{
    this.GetProgram('honeycomb', _VERT_HONEYCOMB, _FRAG_HONEYCOMB);
    this.GetProgram('flip3d', _VERT_3D, _FRAG_TEXTURED);
    this._initHexMeshBuffer('hexTiles', 0.18);
    this._initQuadBuffer3D();
};

CTransitionGL.prototype._renderHoneycomb = function(progress)
{
    let gl = this.gl;
    let prog = this.programs['honeycomb'];
    let flatProg = this.programs['flip3d'];
    if (!prog || !flatProg) return;

    let aspect = this.glCanvas.width / this.glCanvas.height;
    let fov = Math.PI / 4;
    let dist = 1.0 / Math.tan(fov / 2);
    let projection = _Mat4.perspective(fov, aspect, 0.1, 100.0);

    gl.disable(gl.DEPTH_TEST);
    gl.enable(gl.BLEND);
    gl.blendFunc(gl.SRC_ALPHA, gl.ONE_MINUS_SRC_ALPHA);

    // Hex tiles with zoom + Z rotation (background is black via clearColor)
    let zoomAmount = Math.sin(progress * Math.PI) * 1.5;
    let mv = _Mat4.identity();
    mv = _Mat4.translate(mv, 0, 0, -dist + zoomAmount);
    let rotAngle = Math.sin(progress * Math.PI) * 0.15;
    mv = _Mat4.rotateZ(mv, rotAngle);
    gl.useProgram(prog.program);
    gl.uniformMatrix4fv(prog.uniforms['uProjection'], false, projection);
    gl.uniformMatrix4fv(prog.uniforms['uModelView'], false, mv);
    gl.uniform1f(prog.uniforms['uProgress'], progress);

    gl.activeTexture(gl.TEXTURE0);
    gl.bindTexture(gl.TEXTURE_2D, this.textures.slide1);
    gl.uniform1i(prog.uniforms['uTexture1'], 0);

    gl.activeTexture(gl.TEXTURE1);
    gl.bindTexture(gl.TEXTURE_2D, this.textures.slide2);
    gl.uniform1i(prog.uniforms['uTexture2'], 1);

    this._bindHexMesh('hexTiles', prog);
    gl.drawArrays(gl.TRIANGLES, 0, this.buffers['hexTiles'].vertCount);

    // Directional wipe: slide2 sweeps from right to left at the end
    let wipeStart = 0.78;
    if (progress > wipeStart)
    {
        let wipeProgress = (progress - wipeStart) / (1.0 - wipeStart);
        wipeProgress = wipeProgress * wipeProgress * (3.0 - 2.0 * wipeProgress);
        let canvasWidth = this.glCanvas.width;
        let canvasHeight = this.glCanvas.height;
        let cleanWidth = Math.ceil(canvasWidth * wipeProgress);
        let cleanX = canvasWidth - cleanWidth;

        gl.enable(gl.SCISSOR_TEST);
        gl.scissor(cleanX, 0, cleanWidth, canvasHeight);

        gl.useProgram(flatProg.program);
        let mvWipe = _Mat4.translate(_Mat4.identity(), 0, 0, -dist + 0.01);
        gl.uniformMatrix4fv(flatProg.uniforms['uProjection'], false, projection);
        gl.uniformMatrix4fv(flatProg.uniforms['uModelView'], false, mvWipe);
        gl.uniform1f(flatProg.uniforms['uAlpha'], 1.0);
        gl.activeTexture(gl.TEXTURE0);
        gl.bindTexture(gl.TEXTURE_2D, this.textures.slide2);
        gl.uniform1i(flatProg.uniforms['uTexture'], 0);
        this._bindQuad3D(flatProg);
        gl.drawArrays(gl.TRIANGLE_STRIP, 0, 4);

        gl.disable(gl.SCISSOR_TEST);
    }

    gl.disable(gl.BLEND);
};

// ============================================================
// Shaders: Blinds (triangular prism strips)
// ============================================================

let _VERT_BLINDS = [
    'attribute vec3 aPosition;',
    'attribute vec2 aTexCoord;',
    'attribute float aStripIndex;',
    'attribute float aFaceType;',
    'uniform mat4 uProjection;',
    'uniform mat4 uModelView;',
    'uniform float uProgress;',
    'uniform float uStripCount;',
    'uniform float uIsVertical;',
    'uniform float uAspect;',
    'varying vec2 vTexCoord;',
    'varying float vFaceType;',
    'varying float vShade;',
    'void main() {',
    '    vec3 pos = aPosition;',
    '    float idx = aStripIndex;',
    '    float N = uStripCount;',
    '',
    '    float centerIdx = (N - 1.0) * 0.5;',
    '    float d = abs(idx - centerIdx) / max(centerIdx, 1.0);',
    '    float stagger = 0.5;',
    '    float flipDur = 1.0 - stagger;',
    '    float stripStart = d * stagger;',
    '    float sp = clamp((uProgress - stripStart) / flipDur, 0.0, 1.0);',
    '    sp = sp * sp * (3.0 - 2.0 * sp);',
    '',
    '    float angle = sp * 2.0944;',
    '    float c = cos(angle);',
    '    float s = sin(angle);',
    '',
    '    if (uIsVertical > 0.5) {',
    '        float halfW = uAspect / N;',
    '        float centroidZ = -halfW * 1.7321 / 3.0;',
    '        float stripW = 2.0 * uAspect / N;',
    '        float centerX = -uAspect + (idx + 0.5) * stripW;',
    '        float localX = pos.x - centerX;',
    '        float localZ = pos.z - centroidZ;',
    '        pos.x = centerX + localX * c - localZ * s;',
    '        pos.z = centroidZ + localX * s + localZ * c;',
    '    } else {',
    '        float halfH = 1.0 / N;',
    '        float centroidZ = -halfH * 1.7321 / 3.0;',
    '        float stripH = 2.0 / N;',
    '        float centerY = -1.0 + (idx + 0.5) * stripH;',
    '        float localY = pos.y - centerY;',
    '        float localZ = pos.z - centroidZ;',
    '        pos.y = centerY + localY * c - localZ * s;',
    '        pos.z = centroidZ + localY * s + localZ * c;',
    '    }',
    '',
    '    float normalAngle = 1.5708 - aFaceType * 2.0944;',
    '    float normalZ = sin(normalAngle + angle);',
    '    vShade = 0.6 + 0.4 * max(normalZ, 0.0);',
    '',
    '    vTexCoord = aTexCoord;',
    '    vFaceType = aFaceType;',
    '    gl_Position = uProjection * uModelView * vec4(pos, 1.0);',
    '}'
].join('\n');

let _FRAG_BLINDS = [
    'precision mediump float;',
    'uniform sampler2D uTexture1;',
    'uniform sampler2D uTexture2;',
    'varying vec2 vTexCoord;',
    'varying float vFaceType;',
    'varying float vShade;',
    'void main() {',
    '    vec4 color;',
    '    if (vFaceType < 0.5) {',
    '        color = texture2D(uTexture1, vTexCoord);',
    '    } else if (vFaceType < 1.5) {',
    '        color = texture2D(uTexture2, vTexCoord);',
    '    } else {',
    '        color = vec4(0.02, 0.02, 0.02, 1.0);',
    '    }',
    '    gl_FragColor = vec4(color.rgb * vShade, color.a);',
    '}'
].join('\n');

// ============================================================
// Blinds mesh generation (triangular prism strips)
// ============================================================

CTransitionGL.prototype._initBlindsPrismBuffer = function(count, isVertical, name)
{
    if (this.buffers[name]) return;

    let gl = this.gl;
    let aspect = this.glCanvas.width / this.glCanvas.height;
    let hw = aspect, hh = 1.0;
    let SQRT3 = 1.7320508;

    // 3 faces per strip, 4 verts per face = 12 verts per strip
    // 3 faces per strip, 6 indices per face = 18 indices per strip
    let vertCount = count * 12;
    let idxCount = count * 18;
    // position(3) + texcoord(2) + stripIndex(1) + faceType(1) = 7 floats
    let verts = new Float32Array(vertCount * 7);
    let indices = new Uint16Array(idxCount);

    let vi = 0, ii = 0;

    for (let i = 0; i < count; i++)
    {
        let base = i * 12;

        if (isVertical)
        {
            let stripW = 2.0 * hw / count;
            let halfW = stripW / 2.0;
            let depth = halfW * SQRT3;
            let centerX = -hw + (i + 0.5) * stripW;

            // Equilateral triangle cross-section in XZ (front face at z=0):
            let Ax = centerX - halfW, Az = 0;
            let Bx = centerX + halfW, Bz = 0;
            let Cx = centerX,         Cz = -depth;

            let y0 = -hh, y1 = hh;
            let u0 = i / count, u1 = (i + 1) / count;

            // Face 0 (AB): old slide — front face, normal +Z
            verts[vi++]=Ax; verts[vi++]=y0; verts[vi++]=Az;
            verts[vi++]=u0; verts[vi++]=0;  verts[vi++]=i; verts[vi++]=0;
            verts[vi++]=Bx; verts[vi++]=y0; verts[vi++]=Bz;
            verts[vi++]=u1; verts[vi++]=0;  verts[vi++]=i; verts[vi++]=0;
            verts[vi++]=Ax; verts[vi++]=y1; verts[vi++]=Az;
            verts[vi++]=u0; verts[vi++]=1;  verts[vi++]=i; verts[vi++]=0;
            verts[vi++]=Bx; verts[vi++]=y1; verts[vi++]=Bz;
            verts[vi++]=u1; verts[vi++]=1;  verts[vi++]=i; verts[vi++]=0;

            indices[ii++]=base+0; indices[ii++]=base+1; indices[ii++]=base+3;
            indices[ii++]=base+0; indices[ii++]=base+3; indices[ii++]=base+2;

            // Face 1 (BC): new slide — after 120° rotation B→left, C→right
            verts[vi++]=Bx; verts[vi++]=y0; verts[vi++]=Bz;
            verts[vi++]=u0; verts[vi++]=0;  verts[vi++]=i; verts[vi++]=1;
            verts[vi++]=Cx; verts[vi++]=y0; verts[vi++]=Cz;
            verts[vi++]=u1; verts[vi++]=0;  verts[vi++]=i; verts[vi++]=1;
            verts[vi++]=Bx; verts[vi++]=y1; verts[vi++]=Bz;
            verts[vi++]=u0; verts[vi++]=1;  verts[vi++]=i; verts[vi++]=1;
            verts[vi++]=Cx; verts[vi++]=y1; verts[vi++]=Cz;
            verts[vi++]=u1; verts[vi++]=1;  verts[vi++]=i; verts[vi++]=1;

            indices[ii++]=base+4; indices[ii++]=base+5; indices[ii++]=base+7;
            indices[ii++]=base+4; indices[ii++]=base+7; indices[ii++]=base+6;

            // Face 2 (CA): dark face
            verts[vi++]=Cx; verts[vi++]=y0; verts[vi++]=Cz;
            verts[vi++]=0;  verts[vi++]=0;  verts[vi++]=i; verts[vi++]=2;
            verts[vi++]=Ax; verts[vi++]=y0; verts[vi++]=Az;
            verts[vi++]=0;  verts[vi++]=0;  verts[vi++]=i; verts[vi++]=2;
            verts[vi++]=Cx; verts[vi++]=y1; verts[vi++]=Cz;
            verts[vi++]=0;  verts[vi++]=0;  verts[vi++]=i; verts[vi++]=2;
            verts[vi++]=Ax; verts[vi++]=y1; verts[vi++]=Az;
            verts[vi++]=0;  verts[vi++]=0;  verts[vi++]=i; verts[vi++]=2;

            indices[ii++]=base+8;  indices[ii++]=base+9;  indices[ii++]=base+11;
            indices[ii++]=base+8;  indices[ii++]=base+11; indices[ii++]=base+10;
        }
        else
        {
            let stripH = 2.0 * hh / count;
            let halfH = stripH / 2.0;
            let depth = halfH * SQRT3;
            let centerY = -hh + (i + 0.5) * stripH;

            // Equilateral triangle cross-section in YZ (front face at z=0):
            let Ay = centerY - halfH, Az = 0;
            let By = centerY + halfH, Bz = 0;
            let Cy = centerY,         Cz = -depth;

            let x0 = -hw, x1 = hw;
            let v0 = i / count, v1 = (i + 1) / count;

            // Face 0 (AB): old slide — front face, normal +Z
            verts[vi++]=x0; verts[vi++]=Ay; verts[vi++]=Az;
            verts[vi++]=0;  verts[vi++]=v0; verts[vi++]=i; verts[vi++]=0;
            verts[vi++]=x1; verts[vi++]=Ay; verts[vi++]=Az;
            verts[vi++]=1;  verts[vi++]=v0; verts[vi++]=i; verts[vi++]=0;
            verts[vi++]=x0; verts[vi++]=By; verts[vi++]=Bz;
            verts[vi++]=0;  verts[vi++]=v1; verts[vi++]=i; verts[vi++]=0;
            verts[vi++]=x1; verts[vi++]=By; verts[vi++]=Bz;
            verts[vi++]=1;  verts[vi++]=v1; verts[vi++]=i; verts[vi++]=0;

            indices[ii++]=base+0; indices[ii++]=base+1; indices[ii++]=base+3;
            indices[ii++]=base+0; indices[ii++]=base+3; indices[ii++]=base+2;

            // Face 1 (BC): new slide — after 120° rotation B→bottom, C→top
            verts[vi++]=x0; verts[vi++]=By; verts[vi++]=Bz;
            verts[vi++]=0;  verts[vi++]=v0; verts[vi++]=i; verts[vi++]=1;
            verts[vi++]=x1; verts[vi++]=By; verts[vi++]=Bz;
            verts[vi++]=1;  verts[vi++]=v0; verts[vi++]=i; verts[vi++]=1;
            verts[vi++]=x0; verts[vi++]=Cy; verts[vi++]=Cz;
            verts[vi++]=0;  verts[vi++]=v1; verts[vi++]=i; verts[vi++]=1;
            verts[vi++]=x1; verts[vi++]=Cy; verts[vi++]=Cz;
            verts[vi++]=1;  verts[vi++]=v1; verts[vi++]=i; verts[vi++]=1;

            indices[ii++]=base+4; indices[ii++]=base+5; indices[ii++]=base+7;
            indices[ii++]=base+4; indices[ii++]=base+7; indices[ii++]=base+6;

            // Face 2 (CA): dark face
            verts[vi++]=x0; verts[vi++]=Cy; verts[vi++]=Cz;
            verts[vi++]=0;  verts[vi++]=0;  verts[vi++]=i; verts[vi++]=2;
            verts[vi++]=x1; verts[vi++]=Cy; verts[vi++]=Cz;
            verts[vi++]=0;  verts[vi++]=0;  verts[vi++]=i; verts[vi++]=2;
            verts[vi++]=x0; verts[vi++]=Ay; verts[vi++]=Az;
            verts[vi++]=0;  verts[vi++]=0;  verts[vi++]=i; verts[vi++]=2;
            verts[vi++]=x1; verts[vi++]=Ay; verts[vi++]=Az;
            verts[vi++]=0;  verts[vi++]=0;  verts[vi++]=i; verts[vi++]=2;

            indices[ii++]=base+8;  indices[ii++]=base+9;  indices[ii++]=base+11;
            indices[ii++]=base+8;  indices[ii++]=base+11; indices[ii++]=base+10;
        }
    }

    let vbo = gl.createBuffer();
    gl.bindBuffer(gl.ARRAY_BUFFER, vbo);
    gl.bufferData(gl.ARRAY_BUFFER, verts, gl.STATIC_DRAW);

    let ibo = gl.createBuffer();
    gl.bindBuffer(gl.ELEMENT_ARRAY_BUFFER, ibo);
    gl.bufferData(gl.ELEMENT_ARRAY_BUFFER, indices, gl.STATIC_DRAW);

    this.buffers[name] = { vbo: vbo, ibo: ibo, idxCount: idxCount, stride: 28 };
};

CTransitionGL.prototype._bindBlindsPrism = function(name, progInfo)
{
    let gl = this.gl;
    let buf = this.buffers[name];
    if (!buf) return;

    gl.bindBuffer(gl.ARRAY_BUFFER, buf.vbo);
    gl.bindBuffer(gl.ELEMENT_ARRAY_BUFFER, buf.ibo);

    let stride = buf.stride; // 7 floats * 4 = 28
    let aPos  = progInfo.attrs['aPosition'];
    let aTex  = progInfo.attrs['aTexCoord'];
    let aIdx  = progInfo.attrs['aStripIndex'];
    let aFace = progInfo.attrs['aFaceType'];

    if (aPos !== undefined && aPos !== -1)
    {
        gl.enableVertexAttribArray(aPos);
        gl.vertexAttribPointer(aPos, 3, gl.FLOAT, false, stride, 0);
    }
    if (aTex !== undefined && aTex !== -1)
    {
        gl.enableVertexAttribArray(aTex);
        gl.vertexAttribPointer(aTex, 2, gl.FLOAT, false, stride, 12);
    }
    if (aIdx !== undefined && aIdx !== -1)
    {
        gl.enableVertexAttribArray(aIdx);
        gl.vertexAttribPointer(aIdx, 1, gl.FLOAT, false, stride, 20);
    }
    if (aFace !== undefined && aFace !== -1)
    {
        gl.enableVertexAttribArray(aFace);
        gl.vertexAttribPointer(aFace, 1, gl.FLOAT, false, stride, 24);
    }
};

// ============================================================
// Transition: Blinds — triangular prism rotation
// ============================================================

CTransitionGL.prototype._prepareBlinds = function()
{
    this.GetProgram('blindsPrism', _VERT_BLINDS, _FRAG_BLINDS);
    this._initBlindsPrismBuffer(17, true, 'blindsPV');
    this._initBlindsPrismBuffer(14, false, 'blindsPH');
};

CTransitionGL.prototype._renderBlinds = function(progress, param)
{
    let gl = this.gl;
    let prog = this.programs['blindsPrism'];
    if (!prog) return;

    let isVertical = (param === c_oAscSlideTransitionParams.Blinds_Vertical) ? 1.0 : 0.0;
    let meshName = isVertical > 0.5 ? 'blindsPV' : 'blindsPH';
    let stripCount = isVertical > 0.5 ? 17.0 : 14.0;
    let aspect = this.glCanvas.width / this.glCanvas.height;

    let fov = Math.PI / 4;
    let dist = 1.0 / Math.tan(fov / 2);
    let projection = _Mat4.perspective(fov, aspect, 0.1, 100.0);
    let mv = _Mat4.translate(_Mat4.identity(), 0, 0, -dist);

    gl.disable(gl.BLEND);
    gl.enable(gl.DEPTH_TEST);
    gl.enable(gl.CULL_FACE);
    gl.cullFace(gl.BACK);
    gl.frontFace(gl.CCW);

    gl.useProgram(prog.program);
    gl.uniformMatrix4fv(prog.uniforms['uProjection'], false, projection);
    gl.uniformMatrix4fv(prog.uniforms['uModelView'], false, mv);
    gl.uniform1f(prog.uniforms['uProgress'], progress);
    gl.uniform1f(prog.uniforms['uStripCount'], stripCount);
    gl.uniform1f(prog.uniforms['uIsVertical'], isVertical);
    gl.uniform1f(prog.uniforms['uAspect'], aspect);

    gl.activeTexture(gl.TEXTURE0);
    gl.bindTexture(gl.TEXTURE_2D, this.textures.slide1);
    gl.uniform1i(prog.uniforms['uTexture1'], 0);

    gl.activeTexture(gl.TEXTURE1);
    gl.bindTexture(gl.TEXTURE_2D, this.textures.slide2);
    gl.uniform1i(prog.uniforms['uTexture2'], 1);

    this._bindBlindsPrism(meshName, prog);
    gl.drawElements(gl.TRIANGLES, this.buffers[meshName].idxCount, gl.UNSIGNED_SHORT, 0);

    gl.enable(gl.BLEND);
    gl.activeTexture(gl.TEXTURE0);
};

// ============================================================
// Checker shaders — 3D rotating plates in checkerboard pattern
// ============================================================

let _VERT_CHECKER = [
    'attribute vec3 aPosition;',
    'attribute vec2 aTexCoord;',
    'attribute float aCol;',
    'attribute float aRow;',
    'attribute float aFaceType;',
    'uniform mat4 uProjection;',
    'uniform mat4 uModelView;',
    'uniform float uProgress;',
    'uniform float uCols;',
    'uniform float uRows;',
    'uniform float uIsVertical;',
    'uniform float uAspect;',
    'varying vec2 vTexCoord;',
    'varying float vFaceType;',
    'void main() {',
    '    vec3 pos = aPosition;',
    '',
    '    float cellW = 2.0 * uAspect / uCols;',
    '    float cellH = 2.0 / uRows;',
    '    float centerX = -uAspect + (aCol + 0.5) * cellW;',
    '    float centerY = -1.0 + (aRow + 0.5) * cellH;',
    '',
    '    float sweepPos, edgeDist;',
    '    if (uIsVertical > 0.5) {',
    '        sweepPos = aCol / max(uCols - 1.0, 1.0);',
    '        float cr = (uRows - 1.0) * 0.5;',
    '        edgeDist = 1.0 - abs(aRow - cr) / max(cr, 1.0);',
    '    } else {',
    '        sweepPos = 1.0 - aRow / max(uRows - 1.0, 1.0);',
    '        float cc = (uCols - 1.0) * 0.5;',
    '        edgeDist = 1.0 - abs(aCol - cc) / max(cc, 1.0);',
    '    }',
    '    float stagger = sweepPos * 0.5 + edgeDist * 0.15;',
    '    float sp = clamp((uProgress - stagger) / 0.35, 0.0, 1.0);',
    '    sp = sp * sp * (3.0 - 2.0 * sp);',
    '    float angle = sp * 3.14159265;',
    '    float ca = cos(angle);',
    '    float sa = sin(angle);',
    '',
    '    if (uIsVertical > 0.5) {',
    '        float lx = pos.x - centerX;',
    '        pos.x = centerX + lx * ca;',
    '        pos.z = -lx * sa;',
    '    } else {',
    '        float ly = pos.y - centerY;',
    '        pos.y = centerY + ly * ca;',
    '        pos.z = ly * sa;',
    '    }',
    '',
    '    vTexCoord = aTexCoord;',
    '    vFaceType = aFaceType;',
    '    gl_Position = uProjection * uModelView * vec4(pos, 1.0);',
    '}'
].join('\n');

let _FRAG_CHECKER = [
    'precision mediump float;',
    'uniform sampler2D uTexture1;',
    'uniform sampler2D uTexture2;',
    'varying vec2 vTexCoord;',
    'varying float vFaceType;',
    'void main() {',
    '    if (vFaceType < 0.5) {',
    '        gl_FragColor = texture2D(uTexture1, vTexCoord);',
    '    } else {',
    '        gl_FragColor = texture2D(uTexture2, vTexCoord);',
    '    }',
    '}'
].join('\n');

// ============================================================
// Checker mesh generation — two-sided plates with backface culling
// ============================================================

CTransitionGL.prototype._initCheckerBuffer = function(cols, rows, name, isVertical)
{
    if (this.buffers[name]) return;

    let gl = this.gl;
    let aspect = this.glCanvas.width / this.glCanvas.height;
    let hw = aspect, hh = 1.0;
    let cellW = 2.0 * hw / cols;
    let cellH = 2.0 * hh / rows;

    let plateCount = cols * rows;

    // Each plate: front face (4 verts, 6 idx) + back face (4 verts, 6 idx)
    let vertCount = plateCount * 8;
    let idxCount = plateCount * 12;
    // position(3) + texcoord(2) + col(1) + row(1) + faceType(1) = 8 floats
    let verts = new Float32Array(vertCount * 8);
    let indices = new Uint16Array(idxCount);

    let vi = 0, ii = 0, base = 0;

    for (let r = 0; r < rows; r++)
    {
        for (let c = 0; c < cols; c++)
        {
            let x0 = -hw + c * cellW;
            let x1 = x0 + cellW;
            let y0 = -hh + r * cellH;
            let y1 = y0 + cellH;

            let u0 = c / cols, u1 = (c + 1) / cols;
            let v0 = r / rows, v1 = (r + 1) / rows;

            // Front face (CCW from +Z): old slide, faceType=0
            verts[vi++]=x0; verts[vi++]=y0; verts[vi++]=0;
            verts[vi++]=u0; verts[vi++]=v0; verts[vi++]=c; verts[vi++]=r; verts[vi++]=0;
            verts[vi++]=x1; verts[vi++]=y0; verts[vi++]=0;
            verts[vi++]=u1; verts[vi++]=v0; verts[vi++]=c; verts[vi++]=r; verts[vi++]=0;
            verts[vi++]=x0; verts[vi++]=y1; verts[vi++]=0;
            verts[vi++]=u0; verts[vi++]=v1; verts[vi++]=c; verts[vi++]=r; verts[vi++]=0;
            verts[vi++]=x1; verts[vi++]=y1; verts[vi++]=0;
            verts[vi++]=u1; verts[vi++]=v1; verts[vi++]=c; verts[vi++]=r; verts[vi++]=0;

            indices[ii++]=base+0; indices[ii++]=base+1; indices[ii++]=base+3;
            indices[ii++]=base+0; indices[ii++]=base+3; indices[ii++]=base+2;

            // Back face (CCW from -Z after rotation): new slide, faceType=1
            if (isVertical)
            {
                // Y-axis rotation: X swapped for correct back-face winding
                verts[vi++]=x1; verts[vi++]=y0; verts[vi++]=0;
                verts[vi++]=u0; verts[vi++]=v0; verts[vi++]=c; verts[vi++]=r; verts[vi++]=1;
                verts[vi++]=x0; verts[vi++]=y0; verts[vi++]=0;
                verts[vi++]=u1; verts[vi++]=v0; verts[vi++]=c; verts[vi++]=r; verts[vi++]=1;
                verts[vi++]=x1; verts[vi++]=y1; verts[vi++]=0;
                verts[vi++]=u0; verts[vi++]=v1; verts[vi++]=c; verts[vi++]=r; verts[vi++]=1;
                verts[vi++]=x0; verts[vi++]=y1; verts[vi++]=0;
                verts[vi++]=u1; verts[vi++]=v1; verts[vi++]=c; verts[vi++]=r; verts[vi++]=1;
            }
            else
            {
                // X-axis rotation: Y swapped for correct back-face winding
                verts[vi++]=x0; verts[vi++]=y1; verts[vi++]=0;
                verts[vi++]=u0; verts[vi++]=v0; verts[vi++]=c; verts[vi++]=r; verts[vi++]=1;
                verts[vi++]=x1; verts[vi++]=y1; verts[vi++]=0;
                verts[vi++]=u1; verts[vi++]=v0; verts[vi++]=c; verts[vi++]=r; verts[vi++]=1;
                verts[vi++]=x0; verts[vi++]=y0; verts[vi++]=0;
                verts[vi++]=u0; verts[vi++]=v1; verts[vi++]=c; verts[vi++]=r; verts[vi++]=1;
                verts[vi++]=x1; verts[vi++]=y0; verts[vi++]=0;
                verts[vi++]=u1; verts[vi++]=v1; verts[vi++]=c; verts[vi++]=r; verts[vi++]=1;
            }

            indices[ii++]=base+4; indices[ii++]=base+5; indices[ii++]=base+7;
            indices[ii++]=base+4; indices[ii++]=base+7; indices[ii++]=base+6;

            base += 8;
        }
    }

    let vbo = gl.createBuffer();
    gl.bindBuffer(gl.ARRAY_BUFFER, vbo);
    gl.bufferData(gl.ARRAY_BUFFER, verts, gl.STATIC_DRAW);

    let ibo = gl.createBuffer();
    gl.bindBuffer(gl.ELEMENT_ARRAY_BUFFER, ibo);
    gl.bufferData(gl.ELEMENT_ARRAY_BUFFER, indices, gl.STATIC_DRAW);

    this.buffers[name] = { vbo: vbo, ibo: ibo, idxCount: idxCount, stride: 32 };
};

CTransitionGL.prototype._bindChecker = function(name, progInfo)
{
    let gl = this.gl;
    let buf = this.buffers[name];
    if (!buf) return;

    gl.bindBuffer(gl.ARRAY_BUFFER, buf.vbo);
    gl.bindBuffer(gl.ELEMENT_ARRAY_BUFFER, buf.ibo);

    let stride = buf.stride; // 8 floats * 4 = 32
    let aPos  = progInfo.attrs['aPosition'];
    let aTex  = progInfo.attrs['aTexCoord'];
    let aCol  = progInfo.attrs['aCol'];
    let aRow  = progInfo.attrs['aRow'];
    let aFace = progInfo.attrs['aFaceType'];

    if (aPos !== undefined && aPos !== -1)
    {
        gl.enableVertexAttribArray(aPos);
        gl.vertexAttribPointer(aPos, 3, gl.FLOAT, false, stride, 0);
    }
    if (aTex !== undefined && aTex !== -1)
    {
        gl.enableVertexAttribArray(aTex);
        gl.vertexAttribPointer(aTex, 2, gl.FLOAT, false, stride, 12);
    }
    if (aCol !== undefined && aCol !== -1)
    {
        gl.enableVertexAttribArray(aCol);
        gl.vertexAttribPointer(aCol, 1, gl.FLOAT, false, stride, 20);
    }
    if (aRow !== undefined && aRow !== -1)
    {
        gl.enableVertexAttribArray(aRow);
        gl.vertexAttribPointer(aRow, 1, gl.FLOAT, false, stride, 24);
    }
    if (aFace !== undefined && aFace !== -1)
    {
        gl.enableVertexAttribArray(aFace);
        gl.vertexAttribPointer(aFace, 1, gl.FLOAT, false, stride, 28);
    }
};

// ============================================================
// Transition: Checker — 3D rotating plates
// ============================================================

CTransitionGL.prototype._prepareChecker = function()
{
    this.GetProgram('checkerPlate', _VERT_CHECKER, _FRAG_CHECKER);
    this._initCheckerBuffer(7, 5, 'checkerV', true);
    this._initCheckerBuffer(7, 5, 'checkerH', false);
};

CTransitionGL.prototype._renderChecker = function(progress, param)
{
    let gl = this.gl;
    let prog = this.programs['checkerPlate'];
    if (!prog) return;

    // Horizontal param → rotate around Y axis (vertical); Vertical param → rotate around X axis (horizontal)
    let isVertical = (param === c_oAscSlideTransitionParams.Checker_Horizontal) ? 1.0 : 0.0;
    let meshName = isVertical > 0.5 ? 'checkerV' : 'checkerH';
    let aspect = this.glCanvas.width / this.glCanvas.height;

    let fov = Math.PI / 4;
    let dist = 1.0 / Math.tan(fov / 2);
    let projection = _Mat4.perspective(fov, aspect, 0.1, 100.0);
    let mv = _Mat4.translate(_Mat4.identity(), 0, 0, -dist);

    gl.disable(gl.BLEND);
    gl.enable(gl.DEPTH_TEST);
    gl.enable(gl.CULL_FACE);
    gl.cullFace(gl.BACK);
    gl.frontFace(gl.CCW);

    gl.useProgram(prog.program);
    gl.uniformMatrix4fv(prog.uniforms['uProjection'], false, projection);
    gl.uniformMatrix4fv(prog.uniforms['uModelView'], false, mv);
    gl.uniform1f(prog.uniforms['uProgress'], progress);
    gl.uniform1f(prog.uniforms['uCols'], 7.0);
    gl.uniform1f(prog.uniforms['uRows'], 5.0);
    gl.uniform1f(prog.uniforms['uIsVertical'], isVertical);
    gl.uniform1f(prog.uniforms['uAspect'], aspect);

    gl.activeTexture(gl.TEXTURE0);
    gl.bindTexture(gl.TEXTURE_2D, this.textures.slide1);
    gl.uniform1i(prog.uniforms['uTexture1'], 0);

    gl.activeTexture(gl.TEXTURE1);
    gl.bindTexture(gl.TEXTURE_2D, this.textures.slide2);
    gl.uniform1i(prog.uniforms['uTexture2'], 1);

    this._bindChecker(meshName, prog);
    gl.drawElements(gl.TRIANGLES, this.buffers[meshName].idxCount, gl.UNSIGNED_SHORT, 0);

    gl.enable(gl.BLEND);
    gl.activeTexture(gl.TEXTURE0);
};

// ============ Debug: frame capture ============

CTransitionGL.prototype.captureFrames = function(type, param, frameCount)
{
    if (!this.gl || !this.isInitialized) return;

    frameCount = frameCount || 20;

    let w = this.glCanvas.width;
    let h = this.glCanvas.height;
    let thumbW = 400;
    let thumbH = Math.round(thumbW * h / w);
    let labelH = 22;
    let cols = 5;
    let totalFrames = frameCount + 1;
    let rows = Math.ceil(totalFrames / cols);

    let grid = document.createElement('canvas');
    grid.width = cols * thumbW;
    grid.height = rows * (thumbH + labelH);
    let ctx = grid.getContext('2d');
    ctx.fillStyle = '#111';
    ctx.fillRect(0, 0, grid.width, grid.height);

    for (let i = 0; i < totalFrames; i++)
    {
        let progress = i / frameCount;
        this.Render(type, param, progress);

        let col = i % cols;
        let row = Math.floor(i / cols);
        let x = col * thumbW;
        let y = row * (thumbH + labelH);

        ctx.drawImage(this.glCanvas, 0, 0, w, h, x, y + labelH, thumbW, thumbH);
        ctx.fillStyle = '#fff';
        ctx.font = '14px monospace';
        ctx.fillText(Math.round(progress * 100) + '%', x + 4, y + 15);
    }

    grid.toBlob(function(blob)
    {
        let url = URL.createObjectURL(blob);
        let a = document.createElement('a');
        a.href = url;
        a.download = 'transition_' + type + '_' + param + '_frames.png';
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        setTimeout(function() { URL.revokeObjectURL(url); }, 5000);
    }, 'image/png');
};

window['captureTransition'] = function(count)
{
    if (window['_dbgGL'])
    {
        window['_dbgGL'].captureFrames(window['_dbgType'], window['_dbgParam'], count || 20);
    }
    else
    {
        console.log('No active GL transition. Start a transition first.');
    }
};

// Export
window['CTransitionGL'] = CTransitionGL;
