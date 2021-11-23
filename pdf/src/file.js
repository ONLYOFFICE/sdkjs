/*
 * (c) Copyright Ascensio System SIA 2010-2019
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
 * You can contact Ascensio System SIA at 20A-12 Ernesta Birznieka-Upisha
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

(function(window, undefined) {

    var supportImageDataConstructor = (AscCommon.AscBrowser.isIE && !AscCommon.AscBrowser.isIeEdge) ? false : true;
    function CDocMetaSelection()
    {
        this.Page1 = 0;
        this.Line1 = 0;
        this.Glyph1 = 0;

        this.Page2 = 0;
        this.Line2 = 0;
        this.Glyph2 = 0;

        this.IsSelection = false;
    }

    function CFile()
    {
    	this.nativeFile = 0;
    	this.pages = [];
    	this.zoom = 1;
    	this.isUse3d = false;
    	this.cacheManager = null;
    	this.logging = true;
    	this.Selection = new CDocMetaSelection();
    }

    // interface
    CFile.prototype.close = function() 
    {
        if (this.nativeFile)
        {
            this.nativeFile["close"]();
            this.nativeFile = null;
        }
    };
    CFile.prototype.memory = function()
    {
        return this.nativeFile ? this.nativeFile["memory"]() : null;
    };
    CFile.prototype.free = function(pointer)
    {
        this.nativeFile && this.nativeFile["free"](pointer);
    };
    CFile.prototype.getStructure = function() 
    {
        return this.nativeFile ? this.nativeFile["getStructure"]() : [];
    };

	CFile.prototype.getPage = function(pageIndex, width, height, isNoUseCacheManager)
	{
        if (!this.nativeFile)
            return null;
		if (pageIndex < 0 || pageIndex >= this.pages.length)
			return null;

		if (!width) width = this.pages[pageIndex].W;
		if (!height) height = this.pages[pageIndex].H;
		var t0 = performance.now();
		var pixels = this.nativeFile["getPagePixmap"](pageIndex, width, height);
		if (!pixels)
			return null;
        
        var image = null;
		if (!this.logging)
		{
			image = this._pixelsToCanvas(pixels, width, height, isNoUseCacheManager);
		}
		else
		{
			var t1 = performance.now();
			image = this._pixelsToCanvas(pixels, width, height, isNoUseCacheManager);
			var t2 = performance.now();
			//console.log("time: " + (t1 - t0) + ", " + (t2 - t1));
		}
        this.free(pixels);
        return image;
	};

    CFile.prototype.getLinks = function(pageIndex)
    {
        return this.nativeFile ? this.nativeFile["getLinks"](pageIndex) : [];
    };

    CFile.prototype.GetNearestPos = function(pageIndex, x, y)
    {
        var _line = -1;
        var _glyph = -1;
        var minDist = Number.MAX_SAFE_INTEGER;

        if (!this.pages[pageIndex].Lines)
            return { Line : _line, Glyph : _glyph };

        for (let i = 0; i < this.pages[pageIndex].Lines.length; i++)
        {
            let Y = this.pages[pageIndex].Lines[i].Glyphs[0].Y;
            if (Math.abs(Y - y) < minDist)
            {
                minDist = Math.abs(Y - y);
                _line   = i;
            }
        }
        minDist = Number.MAX_SAFE_INTEGER;
        for (let j = 0; j < this.pages[pageIndex].Lines[_line].Glyphs.length; j++)
        {
            let X = this.pages[pageIndex].Lines[_line].Glyphs[j].X;
            if (Math.abs(X - x) < minDist)
            {
                minDist = Math.abs(X - x);
                _glyph  = j;
            }
        }

        return { Line : _line, Glyph : _glyph };
    }

    CFile.prototype.OnUpdateSelection = function(Selection, width, height)
    {
        if (!this.Selection.IsSelection)
            return;

        var sel = this.Selection;
        var page1  = sel.Page1;
        var page2  = sel.Page2;
        var line1  = sel.Line1;
        var line2  = sel.Line2;
        var glyph1 = sel.Glyph1;
        var glyph2 = sel.Glyph2;
        if (page1 == page2 && line1 == line2 && glyph1 == glyph2)
        {
            Selection.IsSelection = false;
            return;
        }
        else if (page1 == page2 && line1 == line2)
        {
            glyph1 = Math.min(sel.Glyph1, sel.Glyph2);
            glyph2 = Math.max(sel.Glyph1, sel.Glyph2);
        }
        else if (page1 == page2)
        {
            if (line1 > line2)
            {
                line1  = sel.Line2;
                line2  = sel.Line1;
                glyph1 = sel.Glyph2;
                glyph2 = sel.Glyph1;
            }
        }

        if (this.cacheManager)
        {
            this.cacheManager.unlock(Selection.Image);
            Selection.Image = this.cacheManager.lock(width, height);
        }
        else
        {
            delete Selection.Image;
            Selection.Image = document.createElement("canvas");
            Selection.Image.width = width;
            Selection.Image.height = height;
        }
        let ctx = Selection.Image.getContext("2d");
        // TODO: После изменения размера экрана наложение областей выделения
        ctx.clearRect(0, 0, Selection.Image.width, Selection.Image.height);
        ctx.globalAlpha = 0.4;
        ctx.fillStyle = "#7070FF";

        for (let page = page1; page <= page2; page++)
        {
            let start = 0;
            let end   = this.pages[page].Lines.length;
            if (page == page1)
                start = line1;
            if (page == page2)
                end = line2;
            for (let line = start; line <= end; line++)
            {
                let Glyphs = this.pages[page].Lines[line].Glyphs;
                let x = Glyphs[0].X;
                let y = Glyphs[0].Y - Glyphs[0].fontSize;
                // первая строка на первой странице
                if (page == page1 && line == start)
                    x = Glyphs[glyph1].X;
                let w = Glyphs[Glyphs.length - 1].X - x;
                let h = Glyphs[0].fontSize;
                // последняя строка на последней странице
                if (page == page2 && line == end)
                    w = Glyphs[glyph2].X - x;

                ctx.fillRect(x, y, w, h);
            }
        }
        ctx.globalAlpha = 1;
        Selection.IsSelection = true;
    }

	CFile.prototype.OnMouseDown = function(pageIndex, Selection, x, y, w, h)
    {
        var ret = this.GetNearestPos(pageIndex, x, y);

        var sel = this.Selection;
        sel.Page1  = pageIndex;
        sel.Line1  = ret.Line;
        sel.Glyph1 = ret.Glyph;

        sel.Page2  = pageIndex;
        sel.Line2  = ret.Line;
        sel.Glyph2 = ret.Glyph;

        sel.IsSelection = true;
        this.OnUpdateSelection(Selection, w, h);
    }

    CFile.prototype.OnMouseMove = function(pageIndex, Selection, x, y, w, h)
    {
        if (false === this.Selection.IsSelection)
            return;
        var ret = this.GetNearestPos(pageIndex, x, y);

        var sel = this.Selection;
        sel.Page2  = pageIndex;
        sel.Line2  = ret.Line;
        sel.Glyph2 = ret.Glyph;

        this.OnUpdateSelection(Selection, w, h);
    }

    CFile.prototype.OnMouseUp = function()
    {
        this.Selection.IsSelection = false;
    }

    CFile.prototype.getPageBase64 = function(pageIndex, width, height)
	{
		var _canvas = this.getPage(pageIndex, width, height);
		if (!_canvas)
			return "";
		
		try
		{
			return _canvas.toDataURL("image/png");
		}
		catch (err)
		{
		}
		
		return "";
	};
	CFile.prototype.isValid = function()
	{
		return this.pages.length > 0;
	};

	// private functions
	CFile.prototype._pixelsToCanvas2d = function(pixels, width, height, isNoUseCacheManager)
    {        
        var canvas = null;
        if (this.cacheManager && isNoUseCacheManager !== true)
        {
            canvas = this.cacheManager.lock(width, height);
        }
        else
        {
            canvas = document.createElement("canvas");
            canvas.width = width;
            canvas.height = height;
        }
        
        var ctx = canvas.getContext("2d");
        var mappedBuffer = new Uint8ClampedArray(this.memory().buffer, pixels, 4 * width * height);
        var imageData = null;
        if (supportImageDataConstructor)
        {
            imageData = new ImageData(mappedBuffer, width, height);
        }
        else
        {
            imageData = ctx.createImageData(width, height);
            imageData.data.set(mappedBuffer, 0);                    
        }
        if (ctx)
            ctx.putImageData(imageData, 0, 0);
        return canvas;
    };

	CFile.prototype._pixelsToCanvas3d = function(pixels, width, height, isNoUseCacheManager) 
    {
        var vs_source = "\
attribute vec2 aVertex;\n\
attribute vec2 aTex;\n\
varying vec2 vTex;\n\
void main() {\n\
	gl_Position = vec4(aVertex, 0.0, 1.0);\n\
	vTex = aTex;\n\
}";

        var fs_source = "\
precision mediump float;\n\
uniform sampler2D uTexture;\n\
varying vec2 vTex;\n\
void main() {\n\
	gl_FragColor = texture2D(uTexture, vTex);\n\
}";
        var canvas = null;
        if (this.cacheManager && isNoUseCacheManager !== true)
        {
            canvas = this.cacheManager.lock(width, height);
        }
        else
        {
            canvas = document.createElement("canvas");
            canvas.width = width;
            canvas.height = height;
        }

        var gl = canvas.getContext('webgl', { preserveDrawingBuffer : true });
        if (!gl)
            throw new Error('FAIL: could not create webgl canvas context');

        var colorCorrect = gl.BROWSER_DEFAULT_WEBGL;
        gl.pixelStorei(gl.UNPACK_COLORSPACE_CONVERSION_WEBGL, colorCorrect);
        gl.pixelStorei(gl.UNPACK_FLIP_Y_WEBGL, true);

        gl.viewport(0, 0, canvas.width, canvas.height);
        gl.clearColor(0, 0, 0, 1);
        gl.clear(gl.COLOR_BUFFER_BIT);

        if (gl.getError() != gl.NONE)
            throw new Error('FAIL: webgl canvas context setup failed');

        function createShader(source, type) {
            var shader = gl.createShader(type);
            gl.shaderSource(shader, source);
            gl.compileShader(shader);
            if (!gl.getShaderParameter(shader, gl.COMPILE_STATUS))
                throw new Error('FAIL: shader ' + id + ' compilation failed');
            return shader;
        }

        var program = gl.createProgram();
        gl.attachShader(program, createShader(vs_source, gl.VERTEX_SHADER));
        gl.attachShader(program, createShader(fs_source, gl.FRAGMENT_SHADER));
        gl.linkProgram(program);
        if (!gl.getProgramParameter(program, gl.LINK_STATUS))
            throw new Error('FAIL: webgl shader program linking failed');
        gl.useProgram(program);

        var texture = gl.createTexture();
        gl.activeTexture(gl.TEXTURE0);
        gl.bindTexture(gl.TEXTURE_2D, texture);
        gl.texParameteri(gl.TEXTURE_2D, gl.TEXTURE_WRAP_S, gl.CLAMP_TO_EDGE);
        gl.texParameteri(gl.TEXTURE_2D, gl.TEXTURE_WRAP_T, gl.CLAMP_TO_EDGE);
        gl.texParameteri(gl.TEXTURE_2D, gl.TEXTURE_MIN_FILTER, gl.LINEAR);
        gl.texParameteri(gl.TEXTURE_2D, gl.TEXTURE_MAG_FILTER, gl.LINEAR);
        gl.texImage2D(gl.TEXTURE_2D, 0, gl.RGBA, width, height, 0, gl.RGBA, gl.UNSIGNED_BYTE, new Uint8Array(this.memory().buffer, pixels, 4 * width * height));

        if (gl.getError() != gl.NONE)
            throw new Error('FAIL: creating webgl image texture failed');

        function createBuffer(data) {
            var buffer = gl.createBuffer();
            gl.bindBuffer(gl.ARRAY_BUFFER, buffer);
            gl.bufferData(gl.ARRAY_BUFFER, data, gl.STATIC_DRAW);
            return buffer;
        }

        var vertexCoords = new Float32Array([-1, 1, -1, -1, 1, -1, 1, 1]);
        var vertexBuffer = createBuffer(vertexCoords);
        var location = gl.getAttribLocation(program, 'aVertex');
        gl.enableVertexAttribArray(location);
        gl.vertexAttribPointer(location, 2, gl.FLOAT, false, 0, 0);

        if (gl.getError() != gl.NONE)
            throw new Error('FAIL: vertex-coord setup failed');

        var texCoords = new Float32Array([0, 1, 0, 0, 1, 0, 1, 1]);
        var texBuffer = createBuffer(texCoords);
        var location = gl.getAttribLocation(program, 'aTex');
        gl.enableVertexAttribArray(location);
        gl.vertexAttribPointer(location, 2, gl.FLOAT, false, 0, 0);

        if (gl.getError() != gl.NONE)
            throw new Error('FAIL: tex-coord setup setup failed');

        gl.drawArrays(gl.TRIANGLE_FAN, 0, 4);
        return canvas;
    };
            
    CFile.prototype._pixelsToCanvas = function(pixels, width, height, isNoUseCacheManager)
    {
        if (!this.isUse3d)
        {
            return this._pixelsToCanvas2d(pixels, width, height, isNoUseCacheManager);
        }

        try
        {
            return this._pixelsToCanvas3d(pixels, width, height, isNoUseCacheManager);
        }
        catch (err)
        {
            this.isUse3d = false;
            if (this.cacheManager)
                this.cacheManager.clear();
            return this._pixelsToCanvas(pixels, width, height, isNoUseCacheManager);
        }
    };

    CFile.prototype.isNeedPassword = function()
    {
        return this.nativeFile ? this.nativeFile["isNeedPassword"]() : false;
    };

    window["AscViewer"] = window["AscViewer"] || {};

    window["AscViewer"]["baseUrl"] = (typeof document !== 'undefined' && document.currentScript) ? "" : "./../src/engine/";
    window["AscViewer"]["baseEngineUrl"] = "./../src/engine/";

    window["AscViewer"].createFile = function(data)
    {
        var file = new CFile();
        file.nativeFile = new window["AscViewer"]["CDrawingFile"]();
        var error = file.nativeFile["loadFromData"](data);
        if (0 === error)
        {
            file.nativeFile.onRepaintPages = function(pages) {
                file.onRepaintPages && file.onRepaintPages(pages);
            };
            file.pages = file.nativeFile["getPages"]();

            for (var i = 0, len = file.pages.length; i < len; i++)
            {
                var page = file.pages[i];
                page.W = page["W"];
                page.H = page["H"];
                page.Dpi = page["Dpi"];
            }

            file.cacheManager = new AscCommon.CCacheManager(); 
            return file;   
        }
        else if (4 === error)
        {
            return file;
        }
        
        file.close();
        return null;
    };

    window["AscViewer"].setFilePassword = function(file, password)
    {
        var error = file.nativeFile["loadFromDataWithPassword"](password);
        if (0 === error)
        {
            file.nativeFile.onRepaintPages = function(pages) {
                file.onRepaintPages && file.onRepaintPages(pages);
            };
            file.pages = file.nativeFile["getPages"]();

            for (var i = 0, len = file.pages.length; i < len; i++)
            {
                var page = file.pages[i];
                page.W = page["W"];
                page.H = page["H"];
                page.Dpi = page["Dpi"];
            }

            file.cacheManager = new AscCommon.CCacheManager();
        }
    };

})(window, undefined);
