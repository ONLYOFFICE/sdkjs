(function(window, undefined) {

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
    CFile.prototype.loadFromData = function(data) {};
    CFile.prototype.load = function(arrayBuffer)
    {
        return this.loadFromData(new Uint8Array(arrayBuffer));
    };

    CFile.prototype.getInfo = function() {};
    CFile.prototype.getPagePixmap = function(pageIndex, width, height) {};
    CFile.prototype.close = function() {};
    CFile.prototype.memory = function() {};
    CFile.prototype.free = function(pointer) {};

	CFile.prototype.getPage = function(pageIndex, width, height, isNoUseCacheManager)
	{
		if (pageIndex < 0 || pageIndex >= this.pages.length)
			return null;
		if (!width) width = this.pages[pageIndex].W;
		if (!height) height = this.pages[pageIndex].H;
		var t0 = performance.now();
		var pixels = this.getPagePixmap(pageIndex, width, height);
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
        
        var mappedBuffer = new Uint8ClampedArray(this.memory().buffer, pixels, 4 * width * height);
        var imageData = new ImageData(mappedBuffer, width, height);
        var ctx = canvas.getContext("2d");
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

    window["AscViewer"] = window["AscViewer"] || {};
    window["AscViewer"].IFile = CFile;

    window["AscViewer"].createFile = function(buffer)
    {
        var data = new Uint8Array(buffer);
        var file = null;

        var maxCheck = Math.min(100, data.length - 5);
        var pdfCheck = [
            "%".charCodeAt(0),
            "P".charCodeAt(0),
            "D".charCodeAt(0),
            "F".charCodeAt(0),
            "-".charCodeAt(0)
        ];
        var isPdf = false;
        for (var i = 0; i < maxCheck; i++)
        {
            if (data[i + 0] == pdfCheck[0] &&
                data[i + 1] == pdfCheck[1] &&
                data[i + 2] == pdfCheck[2] &&
                data[i + 3] == pdfCheck[3] &&
                data[i + 4] == pdfCheck[4])
            {
                isPdf = true;
                break;
            }
        }

        if (isPdf)
        {
            file = new window["AscViewer"].PdfFile();
            file.type = 0;
        }

        if (data.length > 7 && !file &&
            0x41 == data[0] && 0x54 == data[1] && 0x26 == data[2] && 0x54 == data[3] &&
            0x46 == data[4] && 0x4f == data[5] && 0x52 == data[6] && 0x4d == data[7])
        {
            file = new window["AscViewer"].DjVuFile();
        }

        if (!file)
        {
            file = new window["AscViewer"].XpsFile();
        }

        if (!file)
        {
            file = new window["AscViewer"].PdfFile();
            file.type = 1;
        }

        file.loadFromData(data);
        file.cacheManager = new AscCommon.CCacheManager();
        return file;
    };

})(window, undefined);
