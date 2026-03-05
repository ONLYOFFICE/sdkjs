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
        case c_oAscSlideTransitionTypes.FallOver:
            this._prepareFallOver();
            break;
        case c_oAscSlideTransitionTypes.Switch:
            this._prepareSwitch();
            break;
        case c_oAscSlideTransitionTypes.Gallery:
            this._prepareGallery();
            break;
        case c_oAscSlideTransitionTypes.Conveyor:
            this._prepareConveyor();
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
        case c_oAscSlideTransitionTypes.Wind:
            this._prepareWind();
            break;
        case c_oAscSlideTransitionTypes.Drape:
            this._prepareDrape();
            break;
        case c_oAscSlideTransitionTypes.Curtains:
            this._prepareCurtains2();
            break;
        case c_oAscSlideTransitionTypes.Crush:
            this._prepareCrush();
            break;
        case c_oAscSlideTransitionTypes.PeelOff:
            this._preparePeelOff();
            break;
        case c_oAscSlideTransitionTypes.PageCurlSingle:
            this._preparePageCurlSingle();
            break;
        case c_oAscSlideTransitionTypes.PageCurlDouble:
            this._preparePageCurlDouble();
            break;
        case c_oAscSlideTransitionTypes.Vortex:
            this._prepareVortex();
            break;
        case c_oAscSlideTransitionTypes.Prestige:
            this._preparePrestige();
            break;
        case c_oAscSlideTransitionTypes.Fracture:
            this._prepareFracture();
            break;
        case c_oAscSlideTransitionTypes.Airplane:
            this._prepareAirplane();
            break;
        case c_oAscSlideTransitionTypes.Origami:
            this._prepareOrigami();
            break;
        case c_oAscSlideTransitionTypes.Pan:
            this._preparePan();
            break;
        case c_oAscSlideTransitionTypes.Flythrough:
            this._prepareFlythrough();
            break;
        case c_oAscSlideTransitionTypes.Flash:
            this._prepareFlash();
            break;
        case c_oAscSlideTransitionTypes.Reveal:
            this._prepareReveal();
            break;
        case c_oAscSlideTransitionTypes.Honeycomb:
            this._prepareHoneycomb();
            break;
        case c_oAscSlideTransitionTypes.Glitter:
            this._prepareGlitter();
            break;
        case c_oAscSlideTransitionTypes.Shred:
            this._prepareShred();
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
        case c_oAscSlideTransitionTypes.FallOver:
            this._renderFallOver(progress, param);
            break;
        case c_oAscSlideTransitionTypes.Switch:
            this._renderSwitch(progress, param);
            break;
        case c_oAscSlideTransitionTypes.Gallery:
            this._renderGallery(progress, param);
            break;
        case c_oAscSlideTransitionTypes.Conveyor:
            this._renderConveyor(progress, param);
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
        case c_oAscSlideTransitionTypes.Wind:
            this._renderMeshDeform(progress, param, 1, c_oAscSlideTransitionParams.Wind_Default);
            break;
        case c_oAscSlideTransitionTypes.Drape:
            this._renderMeshDeform(progress, param, 2, c_oAscSlideTransitionParams.Drape_Default);
            break;
        case c_oAscSlideTransitionTypes.Curtains:
            this._renderMeshDeform(progress, param, 3, c_oAscSlideTransitionParams.Curtains_Default);
            break;
        case c_oAscSlideTransitionTypes.Crush:
            this._renderMeshDeform(progress, param, 4, c_oAscSlideTransitionParams.Crush_Default);
            break;
        case c_oAscSlideTransitionTypes.PeelOff:
            this._renderPeelOff(progress, param);
            break;
        case c_oAscSlideTransitionTypes.PageCurlSingle:
            this._renderPageCurlSingle(progress, param);
            break;
        case c_oAscSlideTransitionTypes.PageCurlDouble:
            this._renderPageCurlDouble(progress, param);
            break;
        case c_oAscSlideTransitionTypes.Vortex:
            this._renderVortex(progress, param);
            break;
        case c_oAscSlideTransitionTypes.Prestige:
            this._renderScatter(progress, param, 1, c_oAscSlideTransitionParams.Prestige_Default);
            break;
        case c_oAscSlideTransitionTypes.Fracture:
            this._renderScatter(progress, param, 2, c_oAscSlideTransitionParams.Fracture_Default);
            break;
        case c_oAscSlideTransitionTypes.Airplane:
            this._renderFold(progress, param, 4, 'foldAirplane', c_oAscSlideTransitionParams.Airplane_Default);
            break;
        case c_oAscSlideTransitionTypes.Origami:
            this._renderFold(progress, param, 6, 'foldOrigami', c_oAscSlideTransitionParams.Origami_Default);
            break;
        case c_oAscSlideTransitionTypes.Pan:
            this._renderPan(progress, param);
            break;
        case c_oAscSlideTransitionTypes.Flythrough:
            this._renderFlythrough(progress, param);
            break;
        case c_oAscSlideTransitionTypes.Flash:
            this._renderFlash(progress);
            break;
        case c_oAscSlideTransitionTypes.Reveal:
            this._renderReveal(progress, param);
            break;
        case c_oAscSlideTransitionTypes.Honeycomb:
            this._renderHoneycomb(progress);
            break;
        case c_oAscSlideTransitionTypes.Glitter:
            this._renderGlitter(progress, param);
            break;
        case c_oAscSlideTransitionTypes.Shred:
            this._renderShred(progress, param);
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
    let angle = progress * Math.PI / 2; // 0 to 90 degrees
    let sign = isWindow ? -1 : 1;

    // Draw new slide behind (full quad)
    this._initQuadBuffer3D();
    {
        let mv = _Mat4.identity();
        mv = _Mat4.translate(mv, 0, 0, -dist - 0.01);

        gl.uniformMatrix4fv(prog.uniforms['uProjection'], false, projection);
        gl.uniformMatrix4fv(prog.uniforms['uModelView'], false, mv);
        gl.uniform1f(prog.uniforms['uAlpha'], 1.0);

        gl.activeTexture(gl.TEXTURE0);
        gl.bindTexture(gl.TEXTURE_2D, this.textures.slide2);
        gl.uniform1i(prog.uniforms['uTexture'], 0);

        this._bindQuad3D(prog);
        gl.drawArrays(gl.TRIANGLE_STRIP, 0, 4);
    }

    // Draw old slide left/top half
    {
        let hw = aspect;
        let mv = _Mat4.identity();
        mv = _Mat4.translate(mv, 0, 0, -dist);
        if (isHorz)
        {
            // Hinge at top edge, rotate around X
            mv = _Mat4.translate(mv, 0, 1.0, 0);
            mv = _Mat4.rotateX(mv, sign * angle);
            mv = _Mat4.translate(mv, 0, -1.0, 0);
        }
        else
        {
            // Hinge at left edge, rotate around Y
            mv = _Mat4.translate(mv, -hw, 0, 0);
            mv = _Mat4.rotateY(mv, -sign * angle);
            mv = _Mat4.translate(mv, hw, 0, 0);
        }

        gl.uniformMatrix4fv(prog.uniforms['uProjection'], false, projection);
        gl.uniformMatrix4fv(prog.uniforms['uModelView'], false, mv);
        gl.uniform1f(prog.uniforms['uAlpha'], 1.0);

        gl.activeTexture(gl.TEXTURE0);
        gl.bindTexture(gl.TEXTURE_2D, this.textures.slide1);
        gl.uniform1i(prog.uniforms['uTexture'], 0);

        this._bindBuffer3D(isHorz ? 'halfTop' : 'halfLeft', prog);
        gl.drawArrays(gl.TRIANGLE_STRIP, 0, 4);
    }

    // Draw old slide right/bottom half
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
        gl.uniform1f(prog.uniforms['uAlpha'], 1.0);

        gl.activeTexture(gl.TEXTURE0);
        gl.bindTexture(gl.TEXTURE_2D, this.textures.slide1);
        gl.uniform1i(prog.uniforms['uTexture'], 0);

        this._bindBuffer3D(isHorz ? 'halfBottom' : 'halfRight', prog);
        gl.drawArrays(gl.TRIANGLE_STRIP, 0, 4);
    }
};

// ============================================================
// Transition: FallOver — old slide falls forward
// ============================================================

CTransitionGL.prototype._prepareFallOver = function()
{
    this.GetProgram('flip3d', _VERT_3D, _FRAG_TEXTURED);
    this._initQuadBuffer3D();
};

CTransitionGL.prototype._renderFallOver = function(progress, param)
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

    // Determine invX from param offset
    let base = c_oAscSlideTransitionParams.FallOver_Default;
    let offset = param - base;
    let invX = (offset & 1) ? true : false;

    // Draw new slide behind
    {
        let mv = _Mat4.translate(_Mat4.identity(), 0, 0, -dist - 0.01);
        gl.uniformMatrix4fv(prog.uniforms['uProjection'], false, projection);
        gl.uniformMatrix4fv(prog.uniforms['uModelView'], false, mv);
        gl.uniform1f(prog.uniforms['uAlpha'], 1.0);
        gl.activeTexture(gl.TEXTURE0);
        gl.bindTexture(gl.TEXTURE_2D, this.textures.slide2);
        gl.uniform1i(prog.uniforms['uTexture'], 0);
        this._bindQuad3D(prog);
        gl.drawArrays(gl.TRIANGLE_STRIP, 0, 4);
    }

    // Draw old slide falling (rotate around bottom edge)
    {
        let angle = progress * Math.PI / 2; // 0 to 90 degrees
        let mv = _Mat4.identity();
        mv = _Mat4.translate(mv, 0, 0, -dist);
        // Hinge at bottom edge — fall toward viewer (negative X rotation)
        mv = _Mat4.translate(mv, 0, -1.0, 0);
        mv = _Mat4.rotateX(mv, invX ? angle : -angle);
        mv = _Mat4.translate(mv, 0, 1.0, 0);

        let alpha = 1.0 - progress * 0.7;
        gl.uniformMatrix4fv(prog.uniforms['uProjection'], false, projection);
        gl.uniformMatrix4fv(prog.uniforms['uModelView'], false, mv);
        gl.uniform1f(prog.uniforms['uAlpha'], alpha);
        gl.activeTexture(gl.TEXTURE0);
        gl.bindTexture(gl.TEXTURE_2D, this.textures.slide1);
        gl.uniform1i(prog.uniforms['uTexture'], 0);
        this._bindQuad3D(prog);
        gl.drawArrays(gl.TRIANGLE_STRIP, 0, 4);
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
    this._initQuadBuffer3D();
};

CTransitionGL.prototype._renderGallery = function(progress, param)
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

    let isLeft = (param === c_oAscSlideTransitionParams.Gallery_Left);
    let dir = isLeft ? -1 : 1;
    let hw = aspect;

    // Old slide moves off to the side with perspective tilt
    {
        let slideX = dir * progress * hw * 2;
        let tiltAngle = dir * progress * Math.PI / 6; // up to 30 degrees
        let slideZ = -dist - progress * 0.5;

        let mv = _Mat4.identity();
        mv = _Mat4.translate(mv, slideX, 0, slideZ);
        mv = _Mat4.rotateY(mv, tiltAngle);

        gl.uniformMatrix4fv(prog.uniforms['uProjection'], false, projection);
        gl.uniformMatrix4fv(prog.uniforms['uModelView'], false, mv);
        gl.uniform1f(prog.uniforms['uAlpha'], 1.0 - progress * 0.3);
        gl.activeTexture(gl.TEXTURE0);
        gl.bindTexture(gl.TEXTURE_2D, this.textures.slide1);
        gl.uniform1i(prog.uniforms['uTexture'], 0);
        this._bindQuad3D(prog);
        gl.drawArrays(gl.TRIANGLE_STRIP, 0, 4);
    }

    // New slide comes in from the other side
    {
        let slideX = -dir * (1.0 - progress) * hw * 2;
        let tiltAngle = -dir * (1.0 - progress) * Math.PI / 6;
        let slideZ = -dist - (1.0 - progress) * 0.5;

        let mv = _Mat4.identity();
        mv = _Mat4.translate(mv, slideX, 0, slideZ);
        mv = _Mat4.rotateY(mv, tiltAngle);

        gl.uniformMatrix4fv(prog.uniforms['uProjection'], false, projection);
        gl.uniformMatrix4fv(prog.uniforms['uModelView'], false, mv);
        gl.uniform1f(prog.uniforms['uAlpha'], 0.7 + progress * 0.3);
        gl.activeTexture(gl.TEXTURE0);
        gl.bindTexture(gl.TEXTURE_2D, this.textures.slide2);
        gl.uniform1i(prog.uniforms['uTexture'], 0);
        this._bindQuad3D(prog);
        gl.drawArrays(gl.TRIANGLE_STRIP, 0, 4);
    }
};

// ============================================================
// Transition: Conveyor — conveyor belt with tilt at edge
// ============================================================

CTransitionGL.prototype._prepareConveyor = function()
{
    this.GetProgram('flip3d', _VERT_3D, _FRAG_TEXTURED);
    this._initQuadBuffer3D();
};

CTransitionGL.prototype._renderConveyor = function(progress, param)
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

    let isLeft = (param === c_oAscSlideTransitionParams.Conveyor_Left);
    let dir = isLeft ? -1 : 1;
    let hw = aspect;

    // Old slide moves off and tilts at the edge
    {
        let slideX = dir * progress * hw * 2;
        let tiltAngle = progress * Math.PI / 4; // tilt up to 45 degrees around X
        let slideZ = -dist - Math.sin(progress * Math.PI) * 0.3;

        let mv = _Mat4.identity();
        mv = _Mat4.translate(mv, slideX, 0, slideZ);
        mv = _Mat4.rotateX(mv, tiltAngle);

        gl.uniformMatrix4fv(prog.uniforms['uProjection'], false, projection);
        gl.uniformMatrix4fv(prog.uniforms['uModelView'], false, mv);
        gl.uniform1f(prog.uniforms['uAlpha'], 1.0);
        gl.activeTexture(gl.TEXTURE0);
        gl.bindTexture(gl.TEXTURE_2D, this.textures.slide1);
        gl.uniform1i(prog.uniforms['uTexture'], 0);
        this._bindQuad3D(prog);
        gl.drawArrays(gl.TRIANGLE_STRIP, 0, 4);
    }

    // New slide comes in from other side
    {
        let slideX = -dir * (1.0 - progress) * hw * 2;
        let tiltAngle = (1.0 - progress) * Math.PI / 4;
        let slideZ = -dist - Math.sin((1.0 - progress) * Math.PI) * 0.3;

        let mv = _Mat4.identity();
        mv = _Mat4.translate(mv, slideX, 0, slideZ);
        mv = _Mat4.rotateX(mv, -tiltAngle);

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
// Mesh deformation vertex shader (shared by Ripple, Wind, Drape, Curtains, Crush)
// ============================================================

let _VERT_MESH_DEFORM = [
    'attribute vec3 aPosition;',
    'attribute vec2 aTexCoord;',
    'uniform mat4 uProjection;',
    'uniform mat4 uModelView;',
    'uniform float uProgress;',
    'uniform float uTime;',
    'uniform float uDeformType;',
    'uniform float uParam1;',
    'uniform float uParam2;',
    'varying vec2 vTexCoord;',
    'varying float vShade;',
    '',
    'float hash(vec2 p) {',
    '    return fract(sin(dot(p, vec2(12.9898, 78.233))) * 43758.5453);',
    '}',
    '',
    'void main() {',
    '    vec3 pos = aPosition;',
    '    float p = uProgress;',
    '    vShade = 1.0;',
    '',
    '    if (uDeformType < 0.5) {',  // Ripple (type 0) — expanding circle with wave distortion
    '        vec2 diff = pos.xy - vec2(uParam1, uParam2);',
    '        float dist = length(diff);',
    '        float maxR = 4.5;',
    '        float wavefrontR = p * maxR;',
    '        float d = dist - wavefrontR;',
    '',
    '        float envelope = exp(-d * d * 0.8);',
    '        float wave = sin(d * 5.0);',
    '        float waveAmp = 0.30 * envelope;',
    '        pos.z += wave * waveAmp;',
    '',
    '        float behindWave = smoothstep(0.05, -0.3, d);',
    '        pos.z -= behindWave * 0.15;',
    '',
    '        vShade = 1.0 - 0.15 * envelope * (1.0 - wave) * 0.5 * (1.0 - behindWave);',
    '',
    '    } else if (uDeformType < 1.5) {',  // Wind (type 1) — page peeling from edge
    '        float edgeDist = (pos.x * uParam1 + 1.0) * 0.5;',
    '        float sweep = p * 2.0;',
    '        float t = clamp((sweep - edgeDist) * 2.5, 0.0, 1.0);',
    '        float curl = t * 3.14159 * 0.8;',
    '        pos.z += sin(curl) * 0.6 * t;',
    '        pos.x += uParam1 * (1.0 - cos(curl)) * 0.35 * t;',
    '        pos.y += sin(pos.x * 5.0 + uTime * 2.0) * t * 0.06 * uParam2;',
    '        pos.z += sin(pos.y * 6.0 + uTime * 3.0) * t * 0.05;',
    '        vShade = 1.0 - t * 0.35;',
    '',
    '    } else if (uDeformType < 2.5) {',  // Drape (type 2) — old slide sags/drapes
    '        float fromEdge = 1.0 - max(abs(pos.x), abs(pos.y));',
    '        float sagAmount = p * p;',
    '        float sag = sagAmount * (1.0 - fromEdge * 0.7) * 1.8;',
    '        pos.y -= sag * 0.6;',
    '        pos.z -= sag * 0.4;',
    '        float border = 1.0 - smoothstep(0.0, 0.3, fromEdge);',
    '        pos.y -= border * sagAmount * 0.5;',
    '        pos.z -= border * sagAmount * 0.3;',
    '        float wrinkle = sin(pos.x * 15.0 + pos.y * 3.0) * sin(pos.y * 10.0) * sagAmount * 0.03;',
    '        pos.z += wrinkle;',
    '        vShade = 0.75 + 0.25 * (1.0 - sagAmount * 0.6);',
    '',
    '    } else if (uDeformType < 3.5) {',  // Curtains (type 3) — vertical fabric folds pulling apart
    '        float side = sign(pos.x + 0.001);',
    '        float fromCenter = abs(pos.x);',
    '        float pullDist = p * 1.8;',
    '        pos.x += side * pullDist;',
    '',
    '        float foldFreq = 18.0;',
    '        float foldAmp = p * 0.32 * (0.3 + fromCenter * 0.7);',
    '        float foldPhase = fromCenter * foldFreq + pos.y * 0.5;',
    '        pos.z += sin(foldPhase) * foldAmp;',
    '        pos.z += sin(foldPhase * 2.3 + 1.7) * foldAmp * 0.25;',
    '        pos.z += sin(foldPhase * 0.7 + 3.1) * foldAmp * 0.15;',
    '        pos.x += sin(foldPhase * 0.5 + 0.5) * foldAmp * 0.2 * side;',
    '        pos.z -= fromCenter * p * 0.1;',
    '',
    '        float bunching = 1.0 - p * 0.15 * fromCenter;',
    '        pos.y *= bunching;',
    '',
    '        float foldN = sin(foldPhase);',
    '        vShade = 0.55 + 0.45 * abs(foldN);',
    '',
    '    } else {',  // Crush (type 4) — crumple into wrinkled ball
    '        float crushAmt = p * p;',
    '        float contract = 1.0 - crushAmt * 0.82;',
    '        vec2 toCenter = pos.xy * contract;',
    '',
    '        float n1 = hash(aPosition.xy * 5.0);',
    '        float n2 = hash(aPosition.xy * 5.0 + vec2(7.34, 3.21));',
    '        float n3 = hash(aPosition.xy * 5.0 + vec2(13.7, 9.83));',
    '        float n4 = hash(aPosition.xy * 11.0);',
    '        float n5 = hash(aPosition.xy * 11.0 + vec2(5.67, 8.12));',
    '',
    '        float wrinkleAmp = crushAmt * 0.25;',
    '        float fineWrinkle = crushAmt * 0.08;',
    '        pos.x = toCenter.x + (n1 - 0.5) * wrinkleAmp + (n4 - 0.5) * fineWrinkle;',
    '        pos.y = toCenter.y + (n2 - 0.5) * wrinkleAmp + (n5 - 0.5) * fineWrinkle;',
    '',
    '        float dist = length(toCenter);',
    '        float sphereR = 0.4;',
    '        float sphere = sqrt(max(sphereR * sphereR - dist * dist * crushAmt * 1.2, 0.0)) * crushAmt;',
    '        pos.z = -sphere - n3 * wrinkleAmp * 0.4;',
    '        pos.z += sin(aPosition.x * 20.0 + aPosition.y * 15.0) * crushAmt * 0.03;',
    '',
    '        float shade = 0.55 + 0.25 * n3 + 0.2 * (1.0 - crushAmt);',
    '        vShade = shade;',
    '    }',
    '',
    '    vTexCoord = aTexCoord;',
    '    gl_Position = uProjection * uModelView * vec4(pos, 1.0);',
    '}'
].join('\n');

let _FRAG_MESH_DEFORM = [
    'precision mediump float;',
    'uniform sampler2D uTexture;',
    'uniform float uAlpha;',
    'varying vec2 vTexCoord;',
    'varying float vShade;',
    'void main() {',
    '    vec4 color = texture2D(uTexture, vTexCoord);',
    '    gl_FragColor = vec4(color.rgb * vShade, color.a * uAlpha);',
    '}'
].join('\n');

// ============================================================
// Transition: Ripple — wave displacement
// ============================================================

CTransitionGL.prototype._prepareRipple = function()
{
    this.GetProgram('ripple', _VERT_QUAD, _FRAG_RIPPLE);
};

CTransitionGL.prototype._renderMeshDeform = function(progress, param, deformType, baseParam)
{
    let gl = this.gl;
    let meshProg = this.programs['meshDeform'];
    let flatProg = this.programs['flip3d'];
    if (!meshProg || !flatProg) return;

    let aspect = this.glCanvas.width / this.glCanvas.height;
    let fov = Math.PI / 4;
    let dist = 1.0 / Math.tan(fov / 2);
    let projection = _Mat4.perspective(fov, aspect, 0.1, 100.0);

    // Decode invX/invY from param offset
    let offset = param - baseParam;
    let invX = (offset & 1) ? -1.0 : 1.0;
    let invY = (offset & 2) ? -1.0 : 1.0;

    // 1) Draw new slide behind (flat background)
    gl.useProgram(flatProg.program);
    gl.enable(gl.DEPTH_TEST);
    let mvBack = _Mat4.translate(_Mat4.identity(), 0, 0, -dist - 0.02);
    gl.uniformMatrix4fv(flatProg.uniforms['uProjection'], false, projection);
    gl.uniformMatrix4fv(flatProg.uniforms['uModelView'], false, mvBack);
    gl.uniform1f(flatProg.uniforms['uAlpha'], 1.0);
    gl.activeTexture(gl.TEXTURE0);
    gl.bindTexture(gl.TEXTURE_2D, this.textures.slide2);
    gl.uniform1i(flatProg.uniforms['uTexture'], 0);
    this._bindQuad3D(flatProg);
    gl.drawArrays(gl.TRIANGLE_STRIP, 0, 4);

    // 2) Draw old slide with mesh deformation on top
    let mv = _Mat4.translate(_Mat4.identity(), 0, 0, -dist);
    gl.useProgram(meshProg.program);
    gl.uniformMatrix4fv(meshProg.uniforms['uProjection'], false, projection);
    gl.uniformMatrix4fv(meshProg.uniforms['uModelView'], false, mv);
    gl.uniform1f(meshProg.uniforms['uProgress'], progress);
    gl.uniform1f(meshProg.uniforms['uTime'], progress * 6.28);
    gl.uniform1f(meshProg.uniforms['uDeformType'], deformType);
    gl.uniform1f(meshProg.uniforms['uParam1'], invX);
    gl.uniform1f(meshProg.uniforms['uParam2'], invY);

    // Alpha per transition type: curtains/drape stay opaque (reveal by geometry), others fade
    let meshAlpha = 1.0;
    if (deformType === 3) meshAlpha = 1.0;         // Curtains: opaque, reveals by pulling apart
    else if (deformType === 2) meshAlpha = 1.0;     // Drape: opaque, reveals by sagging off
    else if (deformType === 1) meshAlpha = 1.0;     // Wind: opaque, reveals by peeling
    else if (deformType === 4) meshAlpha = Math.max(0.0, 1.0 - progress * progress); // Crush: fade as crumpled
    else meshAlpha = 1.0 - progress * 0.3;          // Ripple: slight fade
    gl.uniform1f(meshProg.uniforms['uAlpha'], meshAlpha);

    gl.activeTexture(gl.TEXTURE0);
    gl.bindTexture(gl.TEXTURE_2D, this.textures.slide1);
    gl.uniform1i(meshProg.uniforms['uTexture'], 0);

    this._bindMesh('mesh80', meshProg);
    gl.drawElements(gl.TRIANGLES, this.buffers['mesh80'].idxCount, gl.UNSIGNED_SHORT, 0);
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

// ============================================================
// Transitions: Wind, Drape, Curtains, Crush — mesh deformation variants
// ============================================================

CTransitionGL.prototype._prepareWind = function()
{
    this.GetProgram('meshDeform', _VERT_MESH_DEFORM, _FRAG_MESH_DEFORM);
    this.GetProgram('flip3d', _VERT_3D, _FRAG_TEXTURED);
    this._initMeshBuffer(80, 80, 'mesh80');
    this._initQuadBuffer3D();
};

CTransitionGL.prototype._prepareDrape = function()
{
    this.GetProgram('meshDeform', _VERT_MESH_DEFORM, _FRAG_MESH_DEFORM);
    this.GetProgram('flip3d', _VERT_3D, _FRAG_TEXTURED);
    this._initMeshBuffer(80, 80, 'mesh80');
    this._initQuadBuffer3D();
};

CTransitionGL.prototype._prepareCurtains2 = function()
{
    this.GetProgram('meshDeform', _VERT_MESH_DEFORM, _FRAG_MESH_DEFORM);
    this.GetProgram('flip3d', _VERT_3D, _FRAG_TEXTURED);
    this._initMeshBuffer(80, 80, 'mesh80');
    this._initQuadBuffer3D();
};

CTransitionGL.prototype._prepareCrush = function()
{
    this.GetProgram('meshDeform', _VERT_MESH_DEFORM, _FRAG_MESH_DEFORM);
    this.GetProgram('flip3d', _VERT_3D, _FRAG_TEXTURED);
    this._initMeshBuffer(80, 80, 'mesh80');
    this._initQuadBuffer3D();
};

// ============================================================
// Transition: Ferris — circular arc rotation
// ============================================================

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

// ============================================================
// Page curl vertex shader (PeelOff, PageCurlSingle, PageCurlDouble)
// ============================================================

let _VERT_PAGE_CURL = [
    'attribute vec3 aPosition;',
    'attribute vec2 aTexCoord;',
    'uniform mat4 uProjection;',
    'uniform mat4 uModelView;',
    'uniform float uProgress;',
    'uniform float uCurlRadius;',
    'uniform float uCurlAngle;',  // sweep direction angle
    'uniform float uCurlPos;',    // current curl position (-1..1)
    'uniform float uFlipX;',      // 1.0 normal, -1.0 mirror X (for double curl)
    'varying vec2 vTexCoord;',
    'varying float vShade;',
    'void main() {',
    '    vec3 pos = aPosition;',
    '    pos.x *= uFlipX;',
    '    float cosA = cos(uCurlAngle);',
    '    float sinA = sin(uCurlAngle);',
    '',
    '    // Project position onto curl axis',
    '    float projected = pos.x * cosA + pos.y * sinA;',
    '    float perpX = pos.x - projected * cosA;',
    '    float perpY = pos.y - projected * sinA;',
    '',
    '    float R = uCurlRadius;',
    '    float threshold = uCurlPos;',
    '    vShade = 1.0;',
    '',
    '    if (projected > threshold) {',
    '        float d = projected - threshold;',
    '        float angle = d / R;',
    '        if (angle > 3.14159) angle = 3.14159;',
    '        float newProj = threshold + sin(angle) * R;',
    '        float z = (1.0 - cos(angle)) * R;',
    '        pos.x = perpX + newProj * cosA;',
    '        pos.y = perpY + newProj * sinA;',
    '        pos.z = z;',
    '        vShade = 0.7 + 0.3 * cos(angle);',
    '    }',
    '',
    '    pos.x *= uFlipX;',
    '    vTexCoord = aTexCoord;',
    '    gl_Position = uProjection * uModelView * vec4(pos, 1.0);',
    '}'
].join('\n');

let _FRAG_PAGE_CURL = [
    'precision mediump float;',
    'uniform sampler2D uTexture;',
    'uniform float uClipSide;',
    'varying vec2 vTexCoord;',
    'varying float vShade;',
    'void main() {',
    '    if (uClipSide > 0.5 && vTexCoord.x < 0.5) discard;',
    '    if (uClipSide < -0.5 && vTexCoord.x > 0.5) discard;',
    '    vec4 color = texture2D(uTexture, vTexCoord);',
    '    gl_FragColor = vec4(color.rgb * vShade, color.a);',
    '}'
].join('\n');

// ============================================================
// Transition: PeelOff — page peels off revealing new slide
// ============================================================

CTransitionGL.prototype._preparePeelOff = function()
{
    this.GetProgram('pageCurl', _VERT_PAGE_CURL, _FRAG_PAGE_CURL);
    this.GetProgram('flip3d', _VERT_3D, _FRAG_TEXTURED);
    this._initMeshBuffer(50, 50, 'mesh50');
    this._initQuadBuffer3D();
};

CTransitionGL.prototype._renderPeelOff = function(progress, param)
{
    let gl = this.gl;
    let curlProg = this.programs['pageCurl'];
    let flatProg = this.programs['flip3d'];
    if (!curlProg || !flatProg) return;

    let aspect = this.glCanvas.width / this.glCanvas.height;
    let fov = Math.PI / 4;
    let dist = 1.0 / Math.tan(fov / 2);
    let projection = _Mat4.perspective(fov, aspect, 0.1, 100.0);
    let mv = _Mat4.translate(_Mat4.identity(), 0, 0, -dist);

    let base = c_oAscSlideTransitionParams.PeelOff_Default;
    let offset = param - base;
    let invX = (offset & 1) ? -1.0 : 1.0;

    // Draw new slide behind (flat)
    gl.useProgram(flatProg.program);
    gl.enable(gl.DEPTH_TEST);
    let mvBack = _Mat4.translate(_Mat4.identity(), 0, 0, -dist - 0.02);
    gl.uniformMatrix4fv(flatProg.uniforms['uProjection'], false, projection);
    gl.uniformMatrix4fv(flatProg.uniforms['uModelView'], false, mvBack);
    gl.uniform1f(flatProg.uniforms['uAlpha'], 1.0);
    gl.activeTexture(gl.TEXTURE0);
    gl.bindTexture(gl.TEXTURE_2D, this.textures.slide2);
    gl.uniform1i(flatProg.uniforms['uTexture'], 0);
    this._bindQuad3D(flatProg);
    gl.drawArrays(gl.TRIANGLE_STRIP, 0, 4);

    // Draw old slide with curl
    gl.useProgram(curlProg.program);
    gl.uniformMatrix4fv(curlProg.uniforms['uProjection'], false, projection);
    gl.uniformMatrix4fv(curlProg.uniforms['uModelView'], false, mv);
    gl.uniform1f(curlProg.uniforms['uProgress'], progress);
    gl.uniform1f(curlProg.uniforms['uCurlRadius'], 0.15);
    gl.uniform1f(curlProg.uniforms['uCurlAngle'], 0.0); // horizontal sweep
    gl.uniform1f(curlProg.uniforms['uFlipX'], 1.0);
    gl.uniform1f(curlProg.uniforms['uClipSide'], 0.0);
    // Curl position sweeps from right to left (or inverted)
    let curlPos = aspect * invX * (1.0 - progress * 2.2);
    gl.uniform1f(curlProg.uniforms['uCurlPos'], curlPos);

    gl.activeTexture(gl.TEXTURE0);
    gl.bindTexture(gl.TEXTURE_2D, this.textures.slide1);
    gl.uniform1i(curlProg.uniforms['uTexture'], 0);

    this._bindMesh('mesh50', curlProg);
    gl.drawElements(gl.TRIANGLES, this.buffers['mesh50'].idxCount, gl.UNSIGNED_SHORT, 0);
};

// ============================================================
// Transition: PageCurlSingle — single page curl from corner
// ============================================================

CTransitionGL.prototype._preparePageCurlSingle = function()
{
    this.GetProgram('pageCurl', _VERT_PAGE_CURL, _FRAG_PAGE_CURL);
    this.GetProgram('flip3d', _VERT_3D, _FRAG_TEXTURED);
    this._initMeshBuffer(50, 50, 'mesh50');
    this._initQuadBuffer3D();
};

CTransitionGL.prototype._renderPageCurlSingle = function(progress, param)
{
    let gl = this.gl;
    let curlProg = this.programs['pageCurl'];
    let flatProg = this.programs['flip3d'];
    if (!curlProg || !flatProg) return;

    let aspect = this.glCanvas.width / this.glCanvas.height;
    let fov = Math.PI / 4;
    let dist = 1.0 / Math.tan(fov / 2);
    let projection = _Mat4.perspective(fov, aspect, 0.1, 100.0);
    let mv = _Mat4.translate(_Mat4.identity(), 0, 0, -dist);

    let base = c_oAscSlideTransitionParams.PageCurlSingle_Default;
    let offset = param - base;
    let invX = (offset & 1) ? true : false;

    // New slide behind
    gl.useProgram(flatProg.program);
    gl.enable(gl.DEPTH_TEST);
    let mvBack = _Mat4.translate(_Mat4.identity(), 0, 0, -dist - 0.02);
    gl.uniformMatrix4fv(flatProg.uniforms['uProjection'], false, projection);
    gl.uniformMatrix4fv(flatProg.uniforms['uModelView'], false, mvBack);
    gl.uniform1f(flatProg.uniforms['uAlpha'], 1.0);
    gl.activeTexture(gl.TEXTURE0);
    gl.bindTexture(gl.TEXTURE_2D, this.textures.slide2);
    gl.uniform1i(flatProg.uniforms['uTexture'], 0);
    this._bindQuad3D(flatProg);
    gl.drawArrays(gl.TRIANGLE_STRIP, 0, 4);

    // Old slide with diagonal curl from corner
    gl.useProgram(curlProg.program);
    gl.uniformMatrix4fv(curlProg.uniforms['uProjection'], false, projection);
    gl.uniformMatrix4fv(curlProg.uniforms['uModelView'], false, mv);
    gl.uniform1f(curlProg.uniforms['uProgress'], progress);
    gl.uniform1f(curlProg.uniforms['uCurlRadius'], 0.12);
    gl.uniform1f(curlProg.uniforms['uFlipX'], 1.0);
    gl.uniform1f(curlProg.uniforms['uClipSide'], 0.0);
    // Diagonal curl angle (~45 degrees from corner)
    let curlAngle = invX ? -Math.PI / 4 : Math.PI / 4;
    gl.uniform1f(curlProg.uniforms['uCurlAngle'], curlAngle);
    let diag = Math.sqrt(aspect * aspect + 1);
    let curlPos = diag * (1.0 - progress * 2.0);
    gl.uniform1f(curlProg.uniforms['uCurlPos'], curlPos);

    gl.activeTexture(gl.TEXTURE0);
    gl.bindTexture(gl.TEXTURE_2D, this.textures.slide1);
    gl.uniform1i(curlProg.uniforms['uTexture'], 0);

    this._bindMesh('mesh50', curlProg);
    gl.drawElements(gl.TRIANGLES, this.buffers['mesh50'].idxCount, gl.UNSIGNED_SHORT, 0);
};

// ============================================================
// Transition: PageCurlDouble — both halves curl simultaneously
// ============================================================

CTransitionGL.prototype._preparePageCurlDouble = function()
{
    this.GetProgram('pageCurl', _VERT_PAGE_CURL, _FRAG_PAGE_CURL);
    this.GetProgram('flip3d', _VERT_3D, _FRAG_TEXTURED);
    this._initMeshBuffer(50, 50, 'mesh50');
    this._initQuadBuffer3D();
};

CTransitionGL.prototype._renderPageCurlDouble = function(progress, param)
{
    let gl = this.gl;
    let curlProg = this.programs['pageCurl'];
    let flatProg = this.programs['flip3d'];
    if (!curlProg || !flatProg) return;

    let aspect = this.glCanvas.width / this.glCanvas.height;
    let fov = Math.PI / 4;
    let dist = 1.0 / Math.tan(fov / 2);
    let projection = _Mat4.perspective(fov, aspect, 0.1, 100.0);
    let mv = _Mat4.translate(_Mat4.identity(), 0, 0, -dist);

    // New slide behind
    gl.useProgram(flatProg.program);
    gl.enable(gl.DEPTH_TEST);
    let mvBack = _Mat4.translate(_Mat4.identity(), 0, 0, -dist - 0.02);
    gl.uniformMatrix4fv(flatProg.uniforms['uProjection'], false, projection);
    gl.uniformMatrix4fv(flatProg.uniforms['uModelView'], false, mvBack);
    gl.uniform1f(flatProg.uniforms['uAlpha'], 1.0);
    gl.activeTexture(gl.TEXTURE0);
    gl.bindTexture(gl.TEXTURE_2D, this.textures.slide2);
    gl.uniform1i(flatProg.uniforms['uTexture'], 0);
    this._bindQuad3D(flatProg);
    gl.drawArrays(gl.TRIANGLE_STRIP, 0, 4);

    // Book-opening: two halves curl outward from center
    let curlPos = aspect * (1.0 - progress * 2.2);

    gl.useProgram(curlProg.program);
    gl.uniformMatrix4fv(curlProg.uniforms['uProjection'], false, projection);
    gl.uniformMatrix4fv(curlProg.uniforms['uModelView'], false, mv);
    gl.uniform1f(curlProg.uniforms['uProgress'], progress);
    gl.uniform1f(curlProg.uniforms['uCurlRadius'], 0.15);
    gl.uniform1f(curlProg.uniforms['uCurlAngle'], 0.0);
    gl.uniform1f(curlProg.uniforms['uCurlPos'], curlPos);

    gl.activeTexture(gl.TEXTURE0);
    gl.bindTexture(gl.TEXTURE_2D, this.textures.slide1);
    gl.uniform1i(curlProg.uniforms['uTexture'], 0);

    this._bindMesh('mesh50', curlProg);

    // Pass 1: right half curls right
    gl.uniform1f(curlProg.uniforms['uFlipX'], 1.0);
    gl.uniform1f(curlProg.uniforms['uClipSide'], 1.0);
    gl.drawElements(gl.TRIANGLES, this.buffers['mesh50'].idxCount, gl.UNSIGNED_SHORT, 0);

    // Pass 2: left half curls left (mirror X)
    gl.uniform1f(curlProg.uniforms['uFlipX'], -1.0);
    gl.uniform1f(curlProg.uniforms['uClipSide'], -1.0);
    gl.drawElements(gl.TRIANGLES, this.buffers['mesh50'].idxCount, gl.UNSIGNED_SHORT, 0);
};

// ============================================================
// Tile-based scatter shader (Vortex, Prestige, Fracture)
// ============================================================

let _VERT_TILE_SCATTER = [
    'attribute vec3 aPosition;',
    'attribute vec2 aTexCoord;',
    'attribute vec3 aTileOffset;',
    'attribute float aTilePhase;',
    'attribute vec2 aTileCenter;',
    'uniform mat4 uProjection;',
    'uniform mat4 uModelView;',
    'uniform float uProgress;',
    'uniform float uScatterType;',
    'varying vec2 vTexCoord;',
    'varying float vAlpha;',
    '',
    'vec3 rotAx(vec3 v, vec3 ax, float a) {',
    '    float c = cos(a); float s = sin(a); float t = 1.0 - c;',
    '    vec3 n = normalize(ax);',
    '    return v * c + cross(n, v) * s + n * dot(n, v) * t;',
    '}',
    '',
    'void main() {',
    '    vec3 pos = aPosition;',
    '    float p = uProgress;',
    '    vec3 local = pos - vec3(aTileCenter, 0.0);',
    '    float t = clamp((p - aTilePhase * 0.5) / 0.5, 0.0, 1.0);',
    '    vAlpha = 1.0 - t;',
    '',
    '    if (uScatterType < 0.5) {',  // Vortex (type 0): spiral outward
    '        vec3 rax = normalize(vec3(aTileOffset.xy, 0.3));',
    '        local = rotAx(local, rax, t * 6.28 * (aTilePhase * 3.0 + 1.5));',
    '        pos = vec3(aTileCenter, 0.0) + local;',
    '        float angle = t * 6.28 * 3.0;',
    '        float radius = t * 2.0;',
    '        pos.x += cos(angle + aTilePhase * 6.28) * radius;',
    '        pos.y += sin(angle + aTilePhase * 6.28) * radius;',
    '        pos.z += t * aTileOffset.z * 3.0;',
    '    } else if (uScatterType < 1.5) {',  // Prestige (type 1): peel from edges
    '        float edgeDist = max(abs(aTileCenter.x), abs(aTileCenter.y));',
    '        float peelT = clamp((t - (1.0 - edgeDist) * 0.3) / 0.7, 0.0, 1.0);',
    '        vec3 rax = normalize(vec3(aTileOffset.x, 0.0, 1.0));',
    '        local = rotAx(local, rax, peelT * 2.5 * sign(aTileOffset.x));',
    '        pos = vec3(aTileCenter, 0.0) + local;',
    '        pos += aTileOffset * peelT * 1.5;',
    '        pos.z += peelT * 1.0;',
    '        vAlpha = 1.0 - peelT;',
    '    } else if (uScatterType < 2.5) {',  // Fracture (type 2): shards crack from bottom
    '        float fromBottom = (aTileCenter.y + 1.0) * 0.5;',
    '        float crackT = clamp((t - fromBottom * 0.4) / 0.6, 0.0, 1.0);',
    '        vec3 rax = normalize(vec3(aTileOffset.x, aTileOffset.z, 0.3));',
    '        local = rotAx(local, rax, crackT * 5.0 * aTilePhase);',
    '        pos = vec3(aTileCenter, 0.0) + local;',
    '        pos.x += aTileOffset.x * crackT * 0.8;',
    '        pos.y -= crackT * crackT * 3.0;',
    '        pos.z += aTileOffset.z * crackT * 2.0 + crackT * 0.5;',
    '        vAlpha = 1.0 - crackT;',
    '    } else {',  // Shred: pieces scatter in random directions with rotation
    '        vec3 rax = normalize(vec3(aTileOffset.xy, 0.5));',
    '        local = rotAx(local, rax, t * 4.0 * (aTilePhase + 0.5));',
    '        pos = vec3(aTileCenter, 0.0) + local;',
    '        pos += aTileOffset * t * 2.0;',
    '        pos.y -= t * t * 1.5;',
    '        pos.z += t * (aTileOffset.z * 1.5 + 0.3);',
    '        vAlpha = 1.0 - t;',
    '    }',
    '',
    '    vTexCoord = aTexCoord;',
    '    gl_Position = uProjection * uModelView * vec4(pos, 1.0);',
    '}'
].join('\n');

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

CTransitionGL.prototype._renderScatter = function(progress, param, scatterType, baseParam)
{
    let gl = this.gl;
    let scatterProg = this.programs['tileScatter'];
    let flatProg = this.programs['flip3d'];
    if (!scatterProg || !flatProg) return;

    let aspect = this.glCanvas.width / this.glCanvas.height;
    let fov = Math.PI / 4;
    let dist = 1.0 / Math.tan(fov / 2);
    let projection = _Mat4.perspective(fov, aspect, 0.1, 100.0);
    let mv = _Mat4.translate(_Mat4.identity(), 0, 0, -dist);

    // Draw new slide behind
    gl.useProgram(flatProg.program);
    gl.enable(gl.DEPTH_TEST);
    let mvBack = _Mat4.translate(_Mat4.identity(), 0, 0, -dist - 0.02);
    gl.uniformMatrix4fv(flatProg.uniforms['uProjection'], false, projection);
    gl.uniformMatrix4fv(flatProg.uniforms['uModelView'], false, mvBack);
    gl.uniform1f(flatProg.uniforms['uAlpha'], 1.0);
    gl.activeTexture(gl.TEXTURE0);
    gl.bindTexture(gl.TEXTURE_2D, this.textures.slide2);
    gl.uniform1i(flatProg.uniforms['uTexture'], 0);
    this._bindQuad3D(flatProg);
    gl.drawArrays(gl.TRIANGLE_STRIP, 0, 4);

    // Draw old slide scattering
    gl.useProgram(scatterProg.program);
    gl.uniformMatrix4fv(scatterProg.uniforms['uProjection'], false, projection);
    gl.uniformMatrix4fv(scatterProg.uniforms['uModelView'], false, mv);
    gl.uniform1f(scatterProg.uniforms['uProgress'], progress);
    gl.uniform1f(scatterProg.uniforms['uScatterType'], scatterType);

    gl.activeTexture(gl.TEXTURE0);
    gl.bindTexture(gl.TEXTURE_2D, this.textures.slide1);
    gl.uniform1i(scatterProg.uniforms['uTexture'], 0);

    let tileName;
    if (scatterType === 0) tileName = 'tilesVortex';
    else if (scatterType === 1) tileName = 'tilesPrestige';
    else tileName = 'tilesFracture';
    this._bindTileMesh(tileName, scatterProg);
    gl.drawArrays(gl.TRIANGLES, 0, this.buffers[tileName].vertCount);
};

// ============================================================
// Transition: Prestige — blocks scatter outward
// ============================================================

CTransitionGL.prototype._preparePrestige = function()
{
    this.GetProgram('tileScatter', _VERT_TILE_SCATTER, _FRAG_TILE_SCATTER);
    this.GetProgram('flip3d', _VERT_3D, _FRAG_TEXTURED);
    this._initTileMeshBuffer(12, 10, 'tilesPrestige');
    this._initQuadBuffer3D();
};

// ============================================================
// Transition: Fracture — shards fall with gravity
// ============================================================

CTransitionGL.prototype._prepareFracture = function()
{
    this.GetProgram('tileScatter', _VERT_TILE_SCATTER, _FRAG_TILE_SCATTER);
    this.GetProgram('flip3d', _VERT_3D, _FRAG_TEXTURED);
    this._initTileMeshBuffer(16, 12, 'tilesFracture');
    this._initQuadBuffer3D();
};

// ============================================================
// Fold vertex shader (Airplane, Origami)
// ============================================================

let _VERT_FOLD = [
    'attribute vec3 aPosition;',
    'attribute vec2 aTexCoord;',
    'attribute float aFoldSegment;',
    'uniform mat4 uProjection;',
    'uniform mat4 uModelView;',
    'uniform float uProgress;',
    'uniform int uFoldCount;',
    'uniform float uInvX;',
    'uniform float uAspect;',
    'varying vec2 vTexCoord;',
    'varying float vShade;',
    'void main() {',
    '    vec3 pos = aPosition;',
    '    float seg = aFoldSegment;',
    '    float p = uProgress;',
    '    vShade = 1.0;',
    '',
    '    float segCount = float(uFoldCount);',
    '    float segDuration = 1.0 / segCount;',
    '    float hw = uAspect;',
    '    float segWidth = 2.0 * hw / segCount;',
    '    float dir = sign(uInvX);',
    '',
    '    // Apply each fold sequentially from first to last',
    '    for (int i = 0; i < 8; i++) {',
    '        if (i >= uFoldCount) break;',
    '        float si = float(i);',
    '        float sp = clamp((p - si * segDuration) / segDuration, 0.0, 1.0);',
    '        if (sp <= 0.0) continue;',
    '        if (seg > si + 0.5) continue;',
    '',
    '        // Crease position for fold i',
    '        float creaseX;',
    '        if (dir > 0.0) {',
    '            creaseX = -hw + (si + 1.0) * segWidth;',
    '        } else {',
    '            creaseX = hw - (si + 1.0) * segWidth;',
    '        }',
    '',
    '        float angle = sp * 3.14159;',
    '        float localX = pos.x - creaseX;',
    '        float c = cos(angle);',
    '        float s = sin(angle);',
    '        float foldDir = mod(si, 2.0) < 0.5 ? 1.0 : -1.0;',
    '        pos.x = creaseX + localX * c;',
    '        pos.z = pos.z * c + foldDir * abs(localX) * s;',
    '',
    '        if (abs(seg - si) < 0.5) {',
    '            vShade = 0.65 + 0.35 * (1.0 - sp * 0.7);',
    '        }',
    '    }',
    '',
    '    vTexCoord = aTexCoord;',
    '    gl_Position = uProjection * uModelView * vec4(pos, 1.0);',
    '}'
].join('\n');

let _FRAG_FOLD = [
    'precision mediump float;',
    'uniform sampler2D uTexture;',
    'varying vec2 vTexCoord;',
    'varying float vShade;',
    'void main() {',
    '    vec4 color = texture2D(uTexture, vTexCoord);',
    '    gl_FragColor = vec4(color.rgb * vShade, color.a);',
    '}'
].join('\n');

// ============================================================
// Fold mesh generation
// ============================================================

CTransitionGL.prototype._initFoldMeshBuffer = function(foldCount, segsPerFold, name)
{
    if (this.buffers[name]) return;

    let gl = this.gl;
    let aspect = this.glCanvas.width / this.glCanvas.height;
    let hw = aspect, hh = 1.0;
    let totalSegsX = foldCount * segsPerFold;
    let segsY = 20;

    let vertCount = (totalSegsX + 1) * (segsY + 1);
    let idxCount = totalSegsX * segsY * 6;
    // position(3) + texcoord(2) + foldSegment(1) = 6 floats
    let verts = new Float32Array(vertCount * 6);
    let indices = new Uint16Array(idxCount);

    let vi = 0;
    for (let iy = 0; iy <= segsY; iy++)
    {
        let v = iy / segsY;
        let y = -hh + v * 2 * hh;
        for (let ix = 0; ix <= totalSegsX; ix++)
        {
            let u = ix / totalSegsX;
            let x = -hw + u * 2 * hw;
            let foldSeg = Math.floor(ix / segsPerFold);
            if (foldSeg >= foldCount) foldSeg = foldCount - 1;
            verts[vi++] = x;
            verts[vi++] = y;
            verts[vi++] = 0;
            verts[vi++] = u;
            verts[vi++] = v;
            verts[vi++] = foldSeg;
        }
    }

    let ii = 0;
    for (let iy = 0; iy < segsY; iy++)
    {
        for (let ix = 0; ix < totalSegsX; ix++)
        {
            let a = iy * (totalSegsX + 1) + ix;
            let b = a + 1;
            let c = a + (totalSegsX + 1);
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

    this.buffers[name] = { vbo: vbo, ibo: ibo, idxCount: idxCount, stride: 24 };
};

CTransitionGL.prototype._bindFoldMesh = function(name, progInfo)
{
    let gl = this.gl;
    let buf = this.buffers[name];
    if (!buf) return;

    gl.bindBuffer(gl.ARRAY_BUFFER, buf.vbo);
    gl.bindBuffer(gl.ELEMENT_ARRAY_BUFFER, buf.ibo);

    let stride = buf.stride; // 6 floats * 4 = 24

    let aPos = progInfo.attrs['aPosition'];
    let aTex = progInfo.attrs['aTexCoord'];
    let aSeg = progInfo.attrs['aFoldSegment'];

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
    if (aSeg !== undefined && aSeg !== -1)
    {
        gl.enableVertexAttribArray(aSeg);
        gl.vertexAttribPointer(aSeg, 1, gl.FLOAT, false, stride, 20);
    }
};

// ============================================================
// Transition: Airplane — sequential fold along crease lines
// ============================================================

CTransitionGL.prototype._prepareAirplane = function()
{
    this.GetProgram('fold', _VERT_FOLD, _FRAG_FOLD);
    this.GetProgram('flip3d', _VERT_3D, _FRAG_TEXTURED);
    this._initFoldMeshBuffer(4, 8, 'foldAirplane');
    this._initQuadBuffer3D();
};

CTransitionGL.prototype._renderFold = function(progress, param, foldCount, meshName, baseParam)
{
    let gl = this.gl;
    let foldProg = this.programs['fold'];
    let flatProg = this.programs['flip3d'];
    if (!foldProg || !flatProg) return;

    let aspect = this.glCanvas.width / this.glCanvas.height;
    let fov = Math.PI / 4;
    let dist = 1.0 / Math.tan(fov / 2);
    let projection = _Mat4.perspective(fov, aspect, 0.1, 100.0);
    let mv = _Mat4.translate(_Mat4.identity(), 0, 0, -dist);

    let offset = param - baseParam;
    let invX = (offset & 1) ? -1.0 : 1.0;

    // New slide behind
    gl.useProgram(flatProg.program);
    gl.enable(gl.DEPTH_TEST);
    let mvBack = _Mat4.translate(_Mat4.identity(), 0, 0, -dist - 0.02);
    gl.uniformMatrix4fv(flatProg.uniforms['uProjection'], false, projection);
    gl.uniformMatrix4fv(flatProg.uniforms['uModelView'], false, mvBack);
    gl.uniform1f(flatProg.uniforms['uAlpha'], 1.0);
    gl.activeTexture(gl.TEXTURE0);
    gl.bindTexture(gl.TEXTURE_2D, this.textures.slide2);
    gl.uniform1i(flatProg.uniforms['uTexture'], 0);
    this._bindQuad3D(flatProg);
    gl.drawArrays(gl.TRIANGLE_STRIP, 0, 4);

    // Old slide with folds
    gl.useProgram(foldProg.program);
    gl.uniformMatrix4fv(foldProg.uniforms['uProjection'], false, projection);
    gl.uniformMatrix4fv(foldProg.uniforms['uModelView'], false, mv);
    gl.uniform1f(foldProg.uniforms['uProgress'], progress);
    gl.uniform1i(foldProg.uniforms['uFoldCount'], foldCount);
    gl.uniform1f(foldProg.uniforms['uInvX'], invX);
    gl.uniform1f(foldProg.uniforms['uAspect'], aspect);

    gl.activeTexture(gl.TEXTURE0);
    gl.bindTexture(gl.TEXTURE_2D, this.textures.slide1);
    gl.uniform1i(foldProg.uniforms['uTexture'], 0);

    this._bindFoldMesh(meshName, foldProg);
    gl.drawElements(gl.TRIANGLES, this.buffers[meshName].idxCount, gl.UNSIGNED_SHORT, 0);
};

// ============================================================
// Transition: Origami — multi-fold crease rotation
// ============================================================

CTransitionGL.prototype._prepareOrigami = function()
{
    this.GetProgram('fold', _VERT_FOLD, _FRAG_FOLD);
    this.GetProgram('flip3d', _VERT_3D, _FRAG_TEXTURED);
    this._initFoldMeshBuffer(6, 6, 'foldOrigami');
    this._initQuadBuffer3D();
};

// ============================================================
// Transition: Pan — 3D perspective slide with depth
// ============================================================

CTransitionGL.prototype._preparePan = function()
{
    this.GetProgram('flip3d', _VERT_3D, _FRAG_TEXTURED);
    this._initQuadBuffer3D();
};

CTransitionGL.prototype._renderPan = function(progress, param)
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
    let hw = aspect;

    let dx = 0, dy = 0;
    switch (param)
    {
        case c_oAscSlideTransitionParams.Pan_Left:  dx = -1; break;
        case c_oAscSlideTransitionParams.Pan_Right: dx = 1;  break;
        case c_oAscSlideTransitionParams.Pan_Up:    dy = 1;  break;
        case c_oAscSlideTransitionParams.Pan_Down:  dy = -1; break;
    }

    let moveX = dx * hw * 2 * progress;
    let moveY = dy * 2 * progress;
    let scaleDown = 1.0 - Math.sin(progress * Math.PI) * 0.08;
    let zBack = -Math.sin(progress * Math.PI) * 0.4;

    // Old slide moves away
    {
        let mv = _Mat4.identity();
        mv = _Mat4.translate(mv, moveX, moveY, -dist + zBack);
        // Apply slight scale-down
        mv[0] *= scaleDown; mv[5] *= scaleDown;
        gl.uniformMatrix4fv(prog.uniforms['uProjection'], false, projection);
        gl.uniformMatrix4fv(prog.uniforms['uModelView'], false, mv);
        gl.uniform1f(prog.uniforms['uAlpha'], 1.0);
        gl.activeTexture(gl.TEXTURE0);
        gl.bindTexture(gl.TEXTURE_2D, this.textures.slide1);
        gl.uniform1i(prog.uniforms['uTexture'], 0);
        this._bindQuad3D(prog);
        gl.drawArrays(gl.TRIANGLE_STRIP, 0, 4);
    }

    // New slide comes in from opposite side
    {
        let oppX = moveX - dx * hw * 2;
        let oppY = moveY - dy * 2;
        let mv = _Mat4.identity();
        mv = _Mat4.translate(mv, oppX, oppY, -dist + zBack);
        mv[0] *= scaleDown; mv[5] *= scaleDown;
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
// Transition: Flythrough — 3D zoom through
// ============================================================

CTransitionGL.prototype._prepareFlythrough = function()
{
    this.GetProgram('flip3d', _VERT_3D, _FRAG_TEXTURED);
    this._initQuadBuffer3D();
};

CTransitionGL.prototype._renderFlythrough = function(progress, param)
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

    let isOut = (param === c_oAscSlideTransitionParams.Flythrough_Out ||
                 param === c_oAscSlideTransitionParams.Flythrough_Out_Bounce);
    let hasBounce = (param === c_oAscSlideTransitionParams.Flythrough_In_Bounce ||
                     param === c_oAscSlideTransitionParams.Flythrough_Out_Bounce);

    let p1 = Math.min(progress * 2.0, 1.0);      // first half: 0→1
    let p2 = Math.max((progress - 0.5) * 2.0, 0); // second half: 0→1

    // Bounce easing for second phase
    let p2eff = p2;
    if (hasBounce && p2 > 0)
    {
        p2eff = p2 < 0.6 ? (p2 / 0.6) * 1.15 :
                1.15 - Math.sin((p2 - 0.6) / 0.4 * Math.PI) * 0.15;
    }

    // Old slide zooms away
    if (p1 < 1.0 || progress < 0.6)
    {
        let zOld = isOut ? -dist + p1 * 3.0 : -dist - p1 * 3.0;
        let alpha = 1.0 - p1;
        let mv = _Mat4.translate(_Mat4.identity(), 0, 0, zOld);
        gl.uniformMatrix4fv(prog.uniforms['uProjection'], false, projection);
        gl.uniformMatrix4fv(prog.uniforms['uModelView'], false, mv);
        gl.uniform1f(prog.uniforms['uAlpha'], Math.max(alpha, 0));
        gl.activeTexture(gl.TEXTURE0);
        gl.bindTexture(gl.TEXTURE_2D, this.textures.slide1);
        gl.uniform1i(prog.uniforms['uTexture'], 0);
        this._bindQuad3D(prog);
        gl.drawArrays(gl.TRIANGLE_STRIP, 0, 4);
    }

    // New slide zooms in
    if (p2 > 0)
    {
        let zNew = isOut ? -dist - (1.0 - p2eff) * 3.0 : -dist + (1.0 - p2eff) * 3.0;
        let alpha = p2;
        let mv = _Mat4.translate(_Mat4.identity(), 0, 0, zNew);
        gl.uniformMatrix4fv(prog.uniforms['uProjection'], false, projection);
        gl.uniformMatrix4fv(prog.uniforms['uModelView'], false, mv);
        gl.uniform1f(prog.uniforms['uAlpha'], Math.min(alpha, 1));
        gl.activeTexture(gl.TEXTURE0);
        gl.bindTexture(gl.TEXTURE_2D, this.textures.slide2);
        gl.uniform1i(prog.uniforms['uTexture'], 0);
        this._bindQuad3D(prog);
        gl.drawArrays(gl.TRIANGLE_STRIP, 0, 4);
    }
};

// ============================================================
// Transition: Flash — white flash between slides (fragment shader)
// ============================================================

let _FRAG_FLASH = [
    'precision mediump float;',
    'uniform sampler2D uTexture1;',
    'uniform sampler2D uTexture2;',
    'uniform float uProgress;',
    'varying vec2 vTexCoord;',
    'void main() {',
    '    vec4 c1 = texture2D(uTexture1, vTexCoord);',
    '    vec4 c2 = texture2D(uTexture2, vTexCoord);',
    '    float flash = 1.0 - abs(uProgress - 0.5) * 2.0;',
    '    flash = flash * flash * flash;',
    '    vec4 slide = uProgress < 0.5 ? c1 : c2;',
    '    gl_FragColor = mix(slide, vec4(1.0), flash);',
    '}'
].join('\n');

CTransitionGL.prototype._prepareFlash = function()
{
    this.GetProgram('flash', _VERT_QUAD, _FRAG_FLASH);
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

CTransitionGL.prototype._renderFlash = function(progress)
{
    let gl = this.gl;
    let prog = this.programs['flash'];
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
// Transition: Reveal — old slide slides away revealing new behind
// ============================================================

CTransitionGL.prototype._prepareReveal = function()
{
    this.GetProgram('flip3d', _VERT_3D, _FRAG_TEXTURED);
    this._initQuadBuffer3D();
    // Create 1x1 black texture for the dark band
    if (!this.textures.black)
    {
        let gl = this.gl;
        let tex = gl.createTexture();
        gl.bindTexture(gl.TEXTURE_2D, tex);
        gl.texImage2D(gl.TEXTURE_2D, 0, gl.RGBA, 1, 1, 0, gl.RGBA, gl.UNSIGNED_BYTE,
            new Uint8Array([0, 0, 0, 255]));
        gl.texParameteri(gl.TEXTURE_2D, gl.TEXTURE_MIN_FILTER, gl.NEAREST);
        gl.texParameteri(gl.TEXTURE_2D, gl.TEXTURE_MAG_FILTER, gl.NEAREST);
        this.textures.black = tex;
    }
};

CTransitionGL.prototype._renderReveal = function(progress, param)
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
    let hw = aspect;

    let isLeft = (param === c_oAscSlideTransitionParams.Reveal_SmoothLeft ||
                  param === c_oAscSlideTransitionParams.Reveal_BlackLeft);
    let isBlack = (param === c_oAscSlideTransitionParams.Reveal_BlackLeft ||
                   param === c_oAscSlideTransitionParams.Reveal_BlackRight);
    let dir = isLeft ? -1 : 1;

    // New slide static behind
    {
        let mv = _Mat4.translate(_Mat4.identity(), 0, 0, -dist - 0.02);
        gl.uniformMatrix4fv(prog.uniforms['uProjection'], false, projection);
        gl.uniformMatrix4fv(prog.uniforms['uModelView'], false, mv);
        gl.uniform1f(prog.uniforms['uAlpha'], 1.0);
        gl.activeTexture(gl.TEXTURE0);
        gl.bindTexture(gl.TEXTURE_2D, this.textures.slide2);
        gl.uniform1i(prog.uniforms['uTexture'], 0);
        this._bindQuad3D(prog);
        gl.drawArrays(gl.TRIANGLE_STRIP, 0, 4);
    }

    // Black band at the leading edge of the sliding old slide
    if (isBlack && this.textures.black)
    {
        let bandW = 0.12;
        let edgeX = dir * progress * hw * 2;
        let bandX = edgeX - dir * bandW * 0.5;
        let mv = _Mat4.translate(_Mat4.identity(), bandX, 0, -dist - 0.01);
        mv[0] *= bandW / hw; // scale X to band width
        gl.uniformMatrix4fv(prog.uniforms['uProjection'], false, projection);
        gl.uniformMatrix4fv(prog.uniforms['uModelView'], false, mv);
        let bandAlpha = Math.sin(progress * Math.PI) * 0.85;
        gl.uniform1f(prog.uniforms['uAlpha'], bandAlpha);
        gl.activeTexture(gl.TEXTURE0);
        gl.bindTexture(gl.TEXTURE_2D, this.textures.black);
        gl.uniform1i(prog.uniforms['uTexture'], 0);
        this._bindQuad3D(prog);
        gl.drawArrays(gl.TRIANGLE_STRIP, 0, 4);
    }

    // Old slide slides away
    {
        let slideX = dir * progress * hw * 2;
        let mv = _Mat4.translate(_Mat4.identity(), slideX, 0, -dist);
        gl.uniformMatrix4fv(prog.uniforms['uProjection'], false, projection);
        gl.uniformMatrix4fv(prog.uniforms['uModelView'], false, mv);
        gl.uniform1f(prog.uniforms['uAlpha'], 1.0);
        gl.activeTexture(gl.TEXTURE0);
        gl.bindTexture(gl.TEXTURE_2D, this.textures.slide1);
        gl.uniform1i(prog.uniforms['uTexture'], 0);
        this._bindQuad3D(prog);
        gl.drawArrays(gl.TRIANGLE_STRIP, 0, 4);
    }
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
// Transition: Glitter — directional tile dissolve with 3D sparkle
// ============================================================

let _VERT_GLITTER = [
    'attribute vec3 aPosition;',
    'attribute vec2 aTexCoord;',
    'attribute vec3 aTileOffset;',
    'attribute float aTilePhase;',
    'attribute vec2 aTileCenter;',
    'uniform mat4 uProjection;',
    'uniform mat4 uModelView;',
    'uniform float uProgress;',
    'uniform float uDirX;',
    'uniform float uDirY;',
    'varying vec2 vTexCoord;',
    'varying float vAlpha;',
    '',
    'void main() {',
    '    vec3 pos = aPosition;',
    '    float p = uProgress;',
    '',
    '    float dirThreshold = (aTileCenter.x * uDirX + aTileCenter.y * uDirY + 2.0) / 4.0;',
    '    float threshold = dirThreshold * 0.65 + aTilePhase * 0.35;',
    '    float reveal = clamp((p - threshold) / 0.12, 0.0, 1.0);',
    '    vAlpha = 1.0 - reveal;',
    '',
    '    // 3D sparkle: tiles near dissolve edge rotate and pop up',
    '    float nearEdge = 1.0 - abs(p - threshold) * 10.0;',
    '    nearEdge = max(nearEdge, 0.0);',
    '    nearEdge = nearEdge * nearEdge;',
    '    vec3 local = pos - vec3(aTileCenter, 0.0);',
    '    float flipAngle = nearEdge * 3.14159 * 0.7;',
    '    float c = cos(flipAngle); float s = sin(flipAngle);',
    '    float axis = aTilePhase * 6.28;',
    '    float nx = cos(axis); float ny = sin(axis);',
    '    float t = 1.0 - c;',
    '    float rx = local.x * (c + nx*nx*t) + local.y * (nx*ny*t) + local.z * (ny*s);',
    '    float ry = local.x * (nx*ny*t) + local.y * (c + ny*ny*t) + local.z * (-nx*s);',
    '    float rz = local.x * (-ny*s) + local.y * (nx*s) + local.z * c;',
    '    pos = vec3(aTileCenter, 0.0) + vec3(rx, ry, rz);',
    '    pos.z += nearEdge * 0.08;',
    '',
    '    vTexCoord = aTexCoord;',
    '    gl_Position = uProjection * uModelView * vec4(pos, 1.0);',
    '}'
].join('\n');

CTransitionGL.prototype._prepareGlitter = function()
{
    this.GetProgram('glitter', _VERT_GLITTER, _FRAG_TILE_SCATTER);
    this.GetProgram('flip3d', _VERT_3D, _FRAG_TEXTURED);
    this._initTileMeshBuffer(45, 30, 'tilesGlitter');
    this._initQuadBuffer3D();
};

CTransitionGL.prototype._renderGlitter = function(progress, param)
{
    let gl = this.gl;
    let glitterProg = this.programs['glitter'];
    let flatProg = this.programs['flip3d'];
    if (!glitterProg || !flatProg) return;

    let aspect = this.glCanvas.width / this.glCanvas.height;
    let fov = Math.PI / 4;
    let dist = 1.0 / Math.tan(fov / 2);
    let projection = _Mat4.perspective(fov, aspect, 0.1, 100.0);

    // Determine direction from param
    let dirX = 0, dirY = 0;
    let base = c_oAscSlideTransitionParams.Glitter_Left_Diamond;
    let offset = param - base;
    let dirIdx = offset % 4;
    switch (dirIdx)
    {
        case 0: dirX = -1; break; // Left
        case 1: dirX = 1;  break; // Right
        case 2: dirY = 1;  break; // Up
        case 3: dirY = -1; break; // Down
    }

    // New slide behind
    gl.useProgram(flatProg.program);
    gl.enable(gl.DEPTH_TEST);
    let mvBack = _Mat4.translate(_Mat4.identity(), 0, 0, -dist - 0.02);
    gl.uniformMatrix4fv(flatProg.uniforms['uProjection'], false, projection);
    gl.uniformMatrix4fv(flatProg.uniforms['uModelView'], false, mvBack);
    gl.uniform1f(flatProg.uniforms['uAlpha'], 1.0);
    gl.activeTexture(gl.TEXTURE0);
    gl.bindTexture(gl.TEXTURE_2D, this.textures.slide2);
    gl.uniform1i(flatProg.uniforms['uTexture'], 0);
    this._bindQuad3D(flatProg);
    gl.drawArrays(gl.TRIANGLE_STRIP, 0, 4);

    // Old slide with glitter tiles
    let mv = _Mat4.translate(_Mat4.identity(), 0, 0, -dist);
    gl.useProgram(glitterProg.program);
    gl.uniformMatrix4fv(glitterProg.uniforms['uProjection'], false, projection);
    gl.uniformMatrix4fv(glitterProg.uniforms['uModelView'], false, mv);
    gl.uniform1f(glitterProg.uniforms['uProgress'], progress);
    gl.uniform1f(glitterProg.uniforms['uDirX'], dirX);
    gl.uniform1f(glitterProg.uniforms['uDirY'], dirY);

    gl.activeTexture(gl.TEXTURE0);
    gl.bindTexture(gl.TEXTURE_2D, this.textures.slide1);
    gl.uniform1i(glitterProg.uniforms['uTexture'], 0);

    this._bindTileMesh('tilesGlitter', glitterProg);
    gl.drawArrays(gl.TRIANGLES, 0, this.buffers['tilesGlitter'].vertCount);
};

// ============================================================
// Transition: Shred — pieces scatter in 3D
// ============================================================

CTransitionGL.prototype._prepareShred = function()
{
    this.GetProgram('tileScatter', _VERT_TILE_SCATTER, _FRAG_TILE_SCATTER);
    this.GetProgram('flip3d', _VERT_3D, _FRAG_TEXTURED);
    this._initQuadBuffer3D();

    // Strips: thin horizontal tiles; Rectangles: square-ish tiles
    let isStrip = (this.currentTransition.param === c_oAscSlideTransitionParams.Shred_StripIn ||
                   this.currentTransition.param === c_oAscSlideTransitionParams.Shred_StripOut);
    if (isStrip)
        this._initTileMeshBuffer(6, 24, 'tilesShredStrip');
    else
        this._initTileMeshBuffer(12, 10, 'tilesShredRect');
};

CTransitionGL.prototype._renderShred = function(progress, param)
{
    let gl = this.gl;
    let scatterProg = this.programs['tileScatter'];
    let flatProg = this.programs['flip3d'];
    if (!scatterProg || !flatProg) return;

    let aspect = this.glCanvas.width / this.glCanvas.height;
    let fov = Math.PI / 4;
    let dist = 1.0 / Math.tan(fov / 2);
    let projection = _Mat4.perspective(fov, aspect, 0.1, 100.0);

    let isIn = (param === c_oAscSlideTransitionParams.Shred_StripIn ||
                param === c_oAscSlideTransitionParams.Shred_RectangleIn);
    let isStrip = (param === c_oAscSlideTransitionParams.Shred_StripIn ||
                   param === c_oAscSlideTransitionParams.Shred_StripOut);
    let tileName = isStrip ? 'tilesShredStrip' : 'tilesShredRect';

    if (isIn)
    {
        // "In": new slide assembles from scattered pieces
        // Draw old slide behind
        gl.useProgram(flatProg.program);
        gl.enable(gl.DEPTH_TEST);
        let mvBack = _Mat4.translate(_Mat4.identity(), 0, 0, -dist - 0.02);
        gl.uniformMatrix4fv(flatProg.uniforms['uProjection'], false, projection);
        gl.uniformMatrix4fv(flatProg.uniforms['uModelView'], false, mvBack);
        gl.uniform1f(flatProg.uniforms['uAlpha'], 1.0);
        gl.activeTexture(gl.TEXTURE0);
        gl.bindTexture(gl.TEXTURE_2D, this.textures.slide1);
        gl.uniform1i(flatProg.uniforms['uTexture'], 0);
        this._bindQuad3D(flatProg);
        gl.drawArrays(gl.TRIANGLE_STRIP, 0, 4);

        // New slide pieces assemble (reverse scatter)
        let mv = _Mat4.translate(_Mat4.identity(), 0, 0, -dist);
        gl.useProgram(scatterProg.program);
        gl.uniformMatrix4fv(scatterProg.uniforms['uProjection'], false, projection);
        gl.uniformMatrix4fv(scatterProg.uniforms['uModelView'], false, mv);
        gl.uniform1f(scatterProg.uniforms['uProgress'], 1.0 - progress); // reverse
        gl.uniform1f(scatterProg.uniforms['uScatterType'], 3);
        gl.activeTexture(gl.TEXTURE0);
        gl.bindTexture(gl.TEXTURE_2D, this.textures.slide2);
        gl.uniform1i(scatterProg.uniforms['uTexture'], 0);
        this._bindTileMesh(tileName, scatterProg);
        gl.drawArrays(gl.TRIANGLES, 0, this.buffers[tileName].vertCount);
    }
    else
    {
        // "Out": old slide scatters away
        // New slide behind
        gl.useProgram(flatProg.program);
        gl.enable(gl.DEPTH_TEST);
        let mvBack = _Mat4.translate(_Mat4.identity(), 0, 0, -dist - 0.02);
        gl.uniformMatrix4fv(flatProg.uniforms['uProjection'], false, projection);
        gl.uniformMatrix4fv(flatProg.uniforms['uModelView'], false, mvBack);
        gl.uniform1f(flatProg.uniforms['uAlpha'], 1.0);
        gl.activeTexture(gl.TEXTURE0);
        gl.bindTexture(gl.TEXTURE_2D, this.textures.slide2);
        gl.uniform1i(flatProg.uniforms['uTexture'], 0);
        this._bindQuad3D(flatProg);
        gl.drawArrays(gl.TRIANGLE_STRIP, 0, 4);

        // Old slide scatters
        let mv = _Mat4.translate(_Mat4.identity(), 0, 0, -dist);
        gl.useProgram(scatterProg.program);
        gl.uniformMatrix4fv(scatterProg.uniforms['uProjection'], false, projection);
        gl.uniformMatrix4fv(scatterProg.uniforms['uModelView'], false, mv);
        gl.uniform1f(scatterProg.uniforms['uProgress'], progress);
        gl.uniform1f(scatterProg.uniforms['uScatterType'], 3);
        gl.activeTexture(gl.TEXTURE0);
        gl.bindTexture(gl.TEXTURE_2D, this.textures.slide1);
        gl.uniform1i(scatterProg.uniforms['uTexture'], 0);
        this._bindTileMesh(tileName, scatterProg);
        gl.drawArrays(gl.TRIANGLES, 0, this.buffers[tileName].vertCount);
    }
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
