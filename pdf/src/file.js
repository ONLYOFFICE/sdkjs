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

(function(window, undefined) {
    function TextStreamReader(data, size)
    {
        this.data = data;
        this.size = size;
        this.pos = 0;

        this.Seek = function(pos)
        {
            if (pos > this.size)
                return 1;
            this.pos = pos;
            return 0;
        };
        this.Skip = function(skip)
        {
            return this.Seek(this.pos + skip);
        };
        // 1 bytes
        this.GetChar = function()
        {
            return this.data[this.pos++];
        };
        // 4 byte
        this.GetLong = function()
        {
            return (this.data[this.pos++] | this.data[this.pos++] << 8 | this.data[this.pos++] << 16 | this.data[this.pos++] << 24);
        };
        this.GetDouble = function()
        {
            return this.GetLong() / 10000;
        };
    }
    var supportImageDataConstructor = (AscCommon.AscBrowser.isIE && !AscCommon.AscBrowser.isIeEdge) ? false : true;

    function CFile()
    {
    	this.nativeFile = 0;
    	this.pages = [];
        this.originalPagesCount = 0;
    	this.zoom = 1;
    	this.isUse3d = false;
    	this.cacheManager = null;
    	this.logging = true;
        this.type = -1;

    	this.Selection = {
            Page1 : 0,
            Line1 : 0,
            Glyph1 : 0,

            Page2 : 0,
            Line2 : 0,
            Glyph2 : 0,

            quads: [],
            IsSelection : false
        };

        this.viewer = null;

        this.maxCanvasSize = 0;
        if (AscCommon.AscBrowser.isAppleDevices || AscCommon.AscBrowser.isAndroid)
            this.maxCanvasSize = 4096;
    }

    // interface
    CFile.prototype.close = function() 
    {
        if (this.nativeFile)
        {
            this.nativeFile["close"]();
            this.nativeFile = null;
            this.pages = [];
        }
    };
    CFile.prototype.getFileBinary = function()
    {
        return this.nativeFile ? this.nativeFile["getFileBinary"]() : null;
    };
    CFile.prototype.getUint8Array = function(ptr, len)
    {
        return this.nativeFile ? this.nativeFile["getUint8Array"](ptr, len) : null;
    };
    CFile.prototype.getUint8ClampedArray = function(ptr, len)
    {
        return this.nativeFile ? this.nativeFile["getUint8ClampedArray"](ptr, len) : null;
    };
    CFile.prototype.free = function(pointer)
    {
        this.nativeFile && this.nativeFile["free"](pointer);
    };
    CFile.prototype.getStructure = function() 
    {
        return this.nativeFile ? this.nativeFile["getStructure"]() : [];
    };
    CFile.prototype.getDocumentInfo = function()
    {
        return this.nativeFile ? this.nativeFile["getDocumentInfo"]() : null;
    };
    CFile.prototype.isNeedCMap = function()
    {
        return this.nativeFile ? this.nativeFile["isNeedCMap"]() : false;
    };
    CFile.prototype.setCMap = function(data)
    {
        if (this.nativeFile)
            this.nativeFile["setCMap"](data);
    };

    CFile.prototype.getPage = function(pageIndex, width, height, isNoUseCacheManager, backgroundColor)
    {
        if (!this.nativeFile)
            return null;
        if (pageIndex < 0 || pageIndex >= this.pages.length || this.pages[pageIndex].originIndex == undefined)
            return null;

        if (!width) width = this.pages[pageIndex].W;
        if (!height) height = this.pages[pageIndex].H;

        var requestW = width;
        var requestH = height;

        if (this.maxCanvasSize > 0)
        {
            if (width > this.maxCanvasSize || height > this.maxCanvasSize)
            {
                var maxKoef = Math.max(width / this.maxCanvasSize, height / this.maxCanvasSize);
                width = (0.5 + (width / maxKoef)) >> 0;
                height = (0.5 + (height / maxKoef)) >> 0;

                if (width > this.maxCanvasSize) width = this.maxCanvasSize;
                if (height > this.maxCanvasSize) height = this.maxCanvasSize;
            }
        }

        var t0 = performance.now();
        var pixels = this.nativeFile["getPagePixmap"](pageIndex, width, height, backgroundColor);
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

        image.requestWidth = requestW;
        image.requestHeight = requestH;
        return image;
    };
    CFile.prototype.addPage = function(pageIndex, pageObj) {
        return this.nativeFile["addPage"](pageIndex, pageObj);
    };
    CFile.prototype.removePage = function(pageIndex) {
        return this.nativeFile["removePage"](pageIndex);
    };
    CFile.prototype.getPageWidth = function(nPage) {
        return this.pages[nPage].W;
    };
    CFile.prototype.getPageHeight = function(nPage) {
        return this.pages[nPage].H;
    };
    CFile.prototype.getLinks = function(pageIndex)
    {
        return this.nativeFile ? this.nativeFile["getLinks"](pageIndex) : [];
    };

    CFile.prototype.getText = function(pageIndex)
    {
        return this.nativeFile ? this.nativeFile["getGlyphs"](pageIndex) : [];
    };

    CFile.prototype.destroyText = function()
    {
        if (this.nativeFile)
            this.nativeFile["destroyTextInfo"]();
    };

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
        var mappedBuffer = this.getUint8ClampedArray(pixels, 4 * width * height);
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
                throw new Error('FAIL: shader compilation failed');
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
        gl.texImage2D(gl.TEXTURE_2D, 0, gl.RGBA, width, height, 0, gl.RGBA, gl.UNSIGNED_BYTE, this.getUint8Array(pixels, 4 * width * height));

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

    // TEXT
    CFile.prototype.getPageTextStream = function(pageIndex) {
        let textCommands = this.pages[pageIndex].text;
        if (!textCommands || 0 === textCommands.length)
            return null;

        return new TextStreamReader(textCommands, textCommands.length);
    };
    CFile.prototype.removeSelection = function() {
        this.Selection = {
			Page1 : 0,
			Line1 : 0,
			Glyph1 : 0,

			Page2 : 0,
			Line2 : 0,
			Glyph2 : 0,
            quads: [],

			IsSelection : false
		}

        this.cacheSelectionQuads([]);
        this.viewer.getPDFDoc().TextSelectTrackHandler.Update();
    };
    CFile.prototype.isSelectionUse = function() {
        return !(this.Selection.Page1  == this.Selection.Page2 && this.Selection.Glyph1 == this.Selection.Glyph2 && this.Selection.Line1 == this.Selection.Line2);
    };
    CFile.prototype.sortSelection = function() {
        let sel = this.Selection;
        let Page1 = 0;
        let Page2 = 0;
        let Line1 = 0;
        let Line2 = 0;
        let Glyph1 = 0;
        let Glyph2 = 0;

        if (sel.Page2 > sel.Page1)
        {
            Page1 = sel.Page1;
            Page2 = sel.Page2;
            Line1 = sel.Line1;
            Line2 = sel.Line2;
            Glyph1 = sel.Glyph1;
            Glyph2 = sel.Glyph2;
        }
        else if (sel.Page2 < sel.Page1)
        {
            Page1 = sel.Page2;
            Page2 = sel.Page1;
            Line1 = sel.Line2;
            Line2 = sel.Line1;
            Glyph1 = sel.Glyph2;
            Glyph2 = sel.Glyph1;
        }
        else if (sel.Page1 === sel.Page2)
        {
            Page1 = sel.Page1;
            Page2 = sel.Page2;

            if (sel.Line1 < sel.Line2)
            {
                Line1 = sel.Line1;
                Line2 = sel.Line2;
                Glyph1 = sel.Glyph1;
                Glyph2 = sel.Glyph2;
            }
            else if (sel.Line2 < sel.Line1)
            {
                Line1 = sel.Line2;
                Line2 = sel.Line1;
                Glyph1 = sel.Glyph2;
                Glyph2 = sel.Glyph1;
            }
            else
            {
                Line1 = sel.Line1;
                Line2 = sel.Line2;

                if (-1 === sel.Glyph1)
                {
                    Glyph1 = sel.Glyph2;
                    Glyph2 = sel.Glyph1;
                }
                else if (-1 === sel.Glyph2)
                {
                    Glyph1 = sel.Glyph1;
                    Glyph2 = sel.Glyph2;
                }
                else if (sel.Glyph1 < sel.Glyph2)
                {
                    Glyph1 = sel.Glyph1;
                    Glyph2 = sel.Glyph2;
                }
                else
                {
                    Glyph1 = sel.Glyph2;
                    Glyph2 = sel.Glyph1;
                }
            }
        }
        return { Page1: Page1, Page2: Page2, Line1: Line1, Line2: Line2, Glyph1: Glyph1, Glyph2: Glyph2 };
    };
    CFile.prototype.getSelection = function() {
        return this.Selection;
    };
    CFile.prototype.onMouseDown = function(pageIndex, x, y) {
        if (this.pages[pageIndex].isRecognized)
            return;
        
        let ret = this.getNearestPos(pageIndex, x, y);
        let sel = this.Selection;

        sel.Page1  = pageIndex;
        sel.Line1  = ret.Line;
        sel.Glyph1 = ret.Glyph;

        sel.Page2  = pageIndex;
        sel.Line2  = ret.Line;
        sel.Glyph2 = ret.Glyph;

        sel.IsSelection = true;
        this.cacheSelectionQuads([]);

        this.onUpdateSelection();
        this.onUpdateOverlay();
    };
    CFile.prototype.onMouseMove = function(pageIndex, x, y) {
        if (false === this.Selection.IsSelection)
            return;

        let ret = this.getNearestPos(pageIndex, x, y);
        let sel = this.Selection;

        sel.Page2  = pageIndex;
        sel.Line2  = ret.Line;
        sel.Glyph2 = ret.Glyph;

        this.onUpdateOverlay();
    };
    CFile.prototype.onMouseUp = function() {
        this.Selection.IsSelection = false;
        this.viewer.getPDFDoc().TextSelectTrackHandler.Update(true);
        this.onUpdateSelection();
        this.onUpdateOverlay();

        if (this.viewer.Api.isMarkerFormat) {
            let oDoc    = this.viewer.getPDFDoc();
            let oViewer = this.viewer;
            let oColor  = oDoc.GetMarkerColor(oViewer.Api.curMarkerType);

            oDoc.DoAction(function() {
                switch (oViewer.Api.curMarkerType) {
                    case AscPDF.ANNOTATIONS_TYPES.Highlight:
                        oViewer.Api.SetHighlight(oColor.r, oColor.g, oColor.b, oColor.a);
                        break;
                    case AscPDF.ANNOTATIONS_TYPES.Underline:
                        oViewer.Api.SetUnderline(oColor.r, oColor.g, oColor.b, oColor.a);
                        break;
                    case AscPDF.ANNOTATIONS_TYPES.Strikeout:
                        oViewer.Api.SetStrikeout(oColor.r, oColor.g, oColor.b, oColor.a);
                        break;
                }
            }, AscDFH.historydescription_Pdf_AddHighlightAnnot);
        }
    };
    CFile.prototype.getNearestPos = function(pageIndex, x, y, bNeedLinePos) {
        let stream = this.getPageTextStream(pageIndex);
        if (!stream) return { Line : -1, Glyph : -1 };

        if (this.type === 2)
        {
            let k = 72 / 96;
            x *= k;
            y *= k;
        }

        let _line  = -1;
        let _glyph = -1;
        let _minDist = Infinity;
        let _minBlockId = -1;
        let _curBlockId = -1;
        let _minBlockDist = Infinity;
        let _predY = Infinity;

        // textline parameters
        let _lineX  = 0;
        let _lineY  = 0;
        let _lineEx = 1;
        let _lineEy = 0;
        let _lineAscent  = 0;
        let _lineDescent = 0;
        let _lineWidth   = 0;
        let _linePrevCharX = 0;
        let _arrayGlyphOffsets = [];
        let _numLine = 0;
        let _linePos = 0;
        let _minLinePos = 0;
        let tmp = 0;

        while (stream.pos < stream.size)
        {
            _lineEx = 1;
            _lineEy = 0;
            _linePrevCharX = 0;
            _arrayGlyphOffsets.splice(0, _arrayGlyphOffsets.length);

            _lineX = stream.GetDouble();
            _lineY = stream.GetDouble();
            if (stream.GetChar())
            {
                _lineEx = stream.GetDouble();
                _lineEy = stream.GetDouble();
            }
            _lineAscent  = stream.GetDouble();
            _lineDescent = stream.GetDouble();
            _lineWidth   = stream.GetDouble();

            if (bNeedLinePos)
                _linePos = stream.pos;

            let nChars = stream.GetLong();
            for (let i = 0; i < nChars; ++i)
            {
                if (i)
                    _linePrevCharX += stream.GetDouble();
                _arrayGlyphOffsets[i] = _linePrevCharX;
                stream.Skip(8);
            }

            if (_lineEx == 1 && _lineEy == 0)
            {
                let _distX = x - _lineX;
                if (y >= (_lineY - _lineAscent) && y <= (_lineY + _lineDescent) && _distX >= 0 && _distX <= _lineWidth)
                { // попали внутрь линии
                    for (_glyph = 1; _glyph < nChars; ++_glyph)
                    { // если символы перекрывают друг друга то текущий выделяется по пересечении начала следующего
                        if (_arrayGlyphOffsets[_glyph] > _distX)
                            break;
                    }

                    let result = {
                        Line: _numLine,
                        Glyph: --_glyph
                    };
                    
                    if (bNeedLinePos) {
                        result.LinePos = _linePos;
                    }
                    
                    return result;
                }

                tmp = Infinity;
                if (y > _lineY - _lineAscent * 2)
                {
                    if (_distX >= 0 && _distX <= _lineWidth)
                        tmp = y - _lineY;
                    else if (_distX < 0)
                        tmp = Math.sqrt((x - _lineX) * (x - _lineX) + (y - _lineY) * (y - _lineY));
                    else if (_distX > _lineWidth)
                    {
                        let _xx1 = _lineX + _lineWidth;
                        tmp = Math.sqrt((x - _xx1) * (x - _xx1) + (y - _lineY) * (y - _lineY));
                    }
                }

                if (_lineY < _predY)
                    _curBlockId++;

                if (tmp < _minBlockDist || (_minBlockId == _curBlockId && y > _lineY && y - _lineY < _minDist))
                {
                    _minDist = y - _lineY;
                    _minBlockDist = tmp;
                    _minBlockId = _curBlockId;
                    _line = _numLine;
                    _minLinePos = _linePos;

                    if (_distX < 0)
                        _glyph = -2;
                    else if (_distX > _lineWidth)
                        _glyph = -1;
                    else
                    {
                        for (_glyph = 1; _glyph < nChars; ++_glyph)
                        {
                            if (_arrayGlyphOffsets[_glyph] > _distX)
                                break;
                        }
                        --_glyph;
                    }
                }
                // Ничего не надо делать, уже найдена более "ближняя" линия
            }
            else
            {
                // определяем точки descent линии
                let _dx = _lineX - _lineEy * _lineDescent;
                let _dy = _lineY + _lineEx * _lineDescent;

                // теперь проекции (со знаком) на линию descent
                let h = (x - _dx) * _lineEy - (y - _dy) * _lineEx;
                let w = (x - _dx) * _lineEx + (y - _dy) * _lineEy;

                if (w >= 0 && w <= _lineWidth && h >= 0 && h <= (_lineDescent + _lineAscent))
                { // попали внутрь линии
                    for (_glyph = 1; _glyph < nChars; ++_glyph)
                    {
                        if (_arrayGlyphOffsets[_glyph] > w)
                            break;
                    }

                    let result = {
                        Line: _numLine,
                        Glyph: --_glyph
                    };
                    
                    if (bNeedLinePos) {
                        result.LinePos = _linePos;
                    }
                    
                    return result;
                }

                if (w >= 0 && w <= _lineWidth)
                    tmp = Math.abs(h - _lineDescent);
                else if (w < 0)
                    tmp = Math.sqrt((x - _lineX) * (x - _lineX) + (y - _lineY) * (y - _lineY));
                else
                {
                    let _tmpX = _lineX + _lineWidth * _lineEx;
                    let _tmpY = _lineY + _lineWidth * _lineEy;
                    tmp = Math.sqrt((x - _tmpX) * (x - _tmpX) + (y - _tmpY) * (y - _tmpY));
                }

                if (tmp < _minBlockDist)
                {
                    _minBlockDist = tmp;
                    _line = _numLine;
                    _minLinePos = _linePos;

                    if (w < 0)
                        _glyph = -2;
                    else if (w > _lineWidth)
                        _glyph = -1;
                    else
                    {
                        for (_glyph = 1; _glyph < nChars; ++_glyph)
                        {
                            if (_arrayGlyphOffsets[_glyph] > w)
                                break;
                        }
                        --_glyph;
                    }
                }
            }
            _predY = _lineY;
            _numLine++;
        }

        let result = {
            Line: _line,
            Glyph: _glyph
        };
        
        if (bNeedLinePos) {
            result.LinePos = _minLinePos;
        }
        
        return result;
    };
    CFile.prototype.selectWholeWord = function(pageIndex, x, y) {
        let ret = this.getNearestPos(pageIndex, x, y, true);
        if (ret.Glyph < 0)
            return;

        let stream = this.getPageTextStream(pageIndex);
        if (!stream) {
            return;
        }

        stream.pos = ret.LinePos;

        let _lineText = "";
        let nChars = stream.GetLong();
        for (let i = 0; i < nChars; ++i)
        {
            if (i)
                stream.Skip(4);
            let nChar = stream.GetLong();
            _lineText += (nChar == 0xFFFF ? " " : String.fromCodePoint(nChar));
            stream.Skip(4);
        }

        let sel = this.Selection;
        sel.Glyph1 = -2;
        sel.Glyph2 = -1;
        sel.IsSelection = true;
        sel.Line1 = ret.Line;
        sel.Line2 = ret.Line;
        sel.Page1 = pageIndex;
        sel.Page2 = pageIndex;
        sel.quads = [];

        let isOnSpace       = false;
        let isOnPunctuation = false;
        if (_lineText[ret.Glyph] == ' ')
            isOnSpace = true;
        else if (undefined != AscCommon.g_aPunctuation[_lineText[ret.Glyph].charCodeAt(0)])
            isOnPunctuation = true;

        if (isOnPunctuation)
            sel.Glyph1 = ret.Glyph;
        else
        {
            for (let i = ret.Glyph - 1; i >=0; --i)
            {
                if (_lineText[i] == ' ' || undefined != AscCommon.g_aPunctuation[_lineText[i].charCodeAt(0)])
                {
                    sel.Glyph1 = i + 1;
                    break;   
                }
            }
        }
        
        if (isOnSpace)
            sel.Glyph2 = ret.Glyph;
        else if (isOnPunctuation && _lineText[ret.Glyph + 1])
            sel.Glyph2 = ret.Glyph + 1;
        else
        {
            for (let i = ret.Glyph + 1; i < _lineText.length; ++i)
            {
                if (_lineText[i] == " " || undefined != AscCommon.g_aPunctuation[_lineText[i].charCodeAt(0)])
                {
                    sel.Glyph2 = i;
                    break;
                }
            }
        }

        this.onUpdateOverlay();
    };
    CFile.prototype.selectWholeRow = function(pageIndex, x, y) {
        let ret = this.getNearestPos(pageIndex, x, y);

        let sel = this.Selection;
        sel.Glyph1 = -2;
        sel.Glyph2 = -1;
        sel.IsSelection = true;
        sel.Line1 = ret.Line;
        sel.Line2 = ret.Line;
        sel.Page1 = pageIndex;
        sel.Page2 = pageIndex;
        sel.quads = [];
        
        this.onUpdateOverlay();
    };
    CFile.prototype.selectWholePage = function(pageIndex) {
        let _numLine = -1;
        let stream = this.getPageTextStream(pageIndex);
        if (!stream) {
            return;
        }
        while (stream.pos < stream.size)
        {
            _numLine++;
            stream.Skip(8);
            if (stream.GetChar())
                stream.Skip(8);
            stream.Skip(12);
            stream.Skip(12 * stream.GetLong() - 4);
        }

        let sel = this.Selection;
        sel.Glyph1 = -2;
        sel.Glyph2 = -1;
        sel.IsSelection = true;
        sel.Line1 = 0;
        sel.Line2 = _numLine + 1;
        sel.Page1 = pageIndex;
        sel.Page2 = pageIndex;
        sel.quads = [];
        
        this.onUpdateOverlay();
    };
    CFile.prototype.selectAll = function() {
        this.removeSelection();
        let sel = this.Selection;
        
        let pagesCount = this.pages.length;
        if (pagesCount)
        {
            let _numLine = -1;
            let stream = this.getPageTextStream(pagesCount - 1);
            if (!stream) {
                return;
            }
            while (stream.pos < stream.size)
            {
                _numLine++;
                stream.Skip(8);
                if (stream.GetChar())
                    stream.Skip(8);
                stream.Skip(12);
                // Не объединять - GetLong прочитает нужное только после skip 12
                stream.Skip(12 * stream.GetLong() - 4);
            }

            sel.Glyph1 = -2;
            sel.Page2 = pagesCount - 1;
            sel.Line2 = _numLine;
            sel.Glyph2 = -1;
        }

        this.onUpdateSelection();
        this.onUpdateOverlay();
        this.viewer.getPDFDoc().TextSelectTrackHandler.Update();
    };
    CFile.prototype.cacheSelectionQuads = function(aQuads) {
        this.Selection.quads = aQuads;
    };
    CFile.prototype.getSelectionQuads = function() {
        let aInfo = [];
        
        if (!this.isSelectionUse())
        {
            this.cacheSelectionQuads(aInfo);
            return aInfo;
        }
        else if (this.Selection.quads.length)
            return this.Selection.quads;
        
        let selection = this.sortSelection();
        let Page1 = selection.Page1;
        let Page2 = selection.Page2;
        let Line1 = selection.Line1;
        let Line2 = selection.Line2;
        let Glyph1 = selection.Glyph1;
        let Glyph2 = selection.Glyph2;

        for (let iPage = Page1; iPage <= Page2; ++iPage)
        {
            let stream = this.getPageTextStream(iPage);
            if (!stream || this.pages[iPage].isRecognized)
                continue;

            let oInfo = { page: iPage, quads: [] };
            let dKoefX = this.pages[iPage].Dpi / 25.4;
            let dKoefY = this.pages[iPage].Dpi / 25.4;

            let startLine = iPage == Page1 ? Line1 : 0;
            let endLine   = iPage == Page2 ? Line2 : Infinity;

            // textline parameters
            let _lineX  = 0;
            let _lineY  = 0;
            let _lineEx = 1;
            let _lineEy = 0;
            let _lineAscent  = 0;
            let _lineDescent = 0;
            let _lineWidth   = 0;
            let _linePrevCharX = 0;
            let _arrayGlyphOffsets = [];
            let iLine = -1;

            while (stream.pos < stream.size)
            {
                iLine++;
                if (iLine < startLine)
                {
                    stream.Skip(8);
                    if (stream.GetChar())
                        stream.Skip(8);
                    stream.Skip(12);
                    stream.Skip(12 * stream.GetLong() - 4);
                    continue;
                }
                if (iLine > endLine)
                    break;

                _lineEx = 1;
                _lineEy = 0;
                _linePrevCharX = 0;
                _arrayGlyphOffsets.splice(0, _arrayGlyphOffsets.length);

                _lineX = stream.GetDouble();
                _lineY = stream.GetDouble();
                if (stream.GetChar())
                {
                    _lineEx = stream.GetDouble();
                    _lineEy = stream.GetDouble();
                }
                _lineAscent  = stream.GetDouble();
                _lineDescent = stream.GetDouble();
                _lineWidth   = stream.GetDouble();

                let nChars = stream.GetLong();
                for (let i = 0; i < nChars; ++i)
                {
                    if (i)
                        _linePrevCharX += stream.GetDouble();
                    _arrayGlyphOffsets[i] = _linePrevCharX;
                    stream.Skip(8);
                }

                let startChar = iPage == Page1 && iLine == Line1 ? Glyph1 : -2;
                let endChar   = iPage == Page2 && iLine == Line2 ? Glyph2 : -1;

                let off1 = 0;
                let off2 = 0;

                if (startChar == -2)
                    off1 = 0;
                else if (startChar == -1)
                    off1 = _lineWidth;
                else
                    off1 = _arrayGlyphOffsets[startChar];

                if (endChar == -2)
                    off2 = 0;
                else if (endChar == -1)
                    off2 = _lineWidth;
                else
                    off2 = _arrayGlyphOffsets[endChar];

                if (off2 <= off1)
                    continue;

                // в принципе код один и тот же. Но почти всегда линии горизонтальные.
                // а для горизонтальной линии все можно пооптимизировать
                if (_lineEx == 1 && _lineEy == 0)
                {
                    let _x = (dKoefX * (_lineX + off1));
                    let _r = (dKoefX * (_lineX + off2));
                    let _y = (dKoefY * (_lineY - _lineAscent));
                    let _b = (dKoefY * (_lineY + _lineDescent));

                    oInfo.quads.push([_x,_y,_r,_y,_x,_b,_r,_b]);
                }
                else
                {
                    // определяем точки descent линии
                    let ortX = -_lineEy;
                    let ortY = _lineEx;

                    let _dx = _lineX + ortX * _lineDescent;
                    let _dy = _lineY + ortY * _lineDescent;

                    let _x1 = _dx + off1 * _lineEx;
                    let _y1 = _dy + off1 * _lineEy;

                    let _x2 = _x1 - ortX * (_lineAscent + _lineDescent);
                    let _y2 = _y1 - ortY * (_lineAscent + _lineDescent);

                    let _x3 = _x2 + (off2 - off1) * _lineEx;
                    let _y3 = _y2 + (off2 - off1) * _lineEy;

                    let _x4 = _x3 + ortX * (_lineAscent + _lineDescent);
                    let _y4 = _y3 + ortY * (_lineAscent + _lineDescent);

                    _x1 = (dKoefX * _x1);
                    _x2 = (dKoefX * _x2);
                    _x3 = (dKoefX * _x3);
                    _x4 = (dKoefX * _x4);

                    _y1 = (dKoefY * _y1);
                    _y2 = (dKoefY * _y2);
                    _y3 = (dKoefY * _y3);
                    _y4 = (dKoefY * _y4);

                    oInfo.quads.push([_x2,_y2, _x3,_y3, _x1,_y1, _x4,_y4]);
                }
            }
            if (oInfo.quads.length > 0)
                aInfo.push(oInfo);
        }
        
        this.cacheSelectionQuads(aInfo);
        return aInfo;
    };
    CFile.prototype.drawSelection = function(pageIndex, overlay, x, y)
    {
        if (this.pages[pageIndex].isRecognized)
            return;
        
        let stream = this.getPageTextStream(pageIndex);
        if (!stream)
            return;

        let selection = this.sortSelection();
        let Page1 = selection.Page1;
        let Page2 = selection.Page2;
        let Line1 = selection.Line1;
        let Line2 = selection.Line2;
        let Glyph1 = selection.Glyph1;
        let Glyph2 = selection.Glyph2;

        if (Page1 > pageIndex || Page2 < pageIndex)
            return;

        // textline parameters
        let _lineX  = 0;
        let _lineY  = 0;
        let _lineEx = 1;
        let _lineEy = 0;
        let _lineAscent  = 0;
        let _lineDescent = 0;
        let _lineWidth   = 0;
        let _linePrevCharX = 0;
        let _arrayGlyphOffsets = [];
        let iLine = -1;

        let width = AscCommon.AscBrowser.convertToRetinaValue(this.viewer.drawingPages[pageIndex].W, true) >> 0;
        let height = AscCommon.AscBrowser.convertToRetinaValue(this.viewer.drawingPages[pageIndex].H, true) >> 0;

        let dKoefX = width  / this.pages[pageIndex].W;
        let dKoefY = height / this.pages[pageIndex].H;
        dKoefX *= (this.pages[pageIndex].Dpi / 25.4);
        dKoefY *= (this.pages[pageIndex].Dpi / 25.4);

        let startLine = pageIndex == Page1 ? Line1 : 0;
        let endLine   = pageIndex == Page2 ? Line2 : Infinity;

        while (stream.pos < stream.size)
        {
            iLine++;
            if (iLine < startLine)
            {
                stream.Skip(8);
                if (stream.GetChar())
                    stream.Skip(8);
                stream.Skip(12);
                stream.Skip(12 * stream.GetLong() - 4);
                continue;
            }
            if (iLine > endLine)
                break;

            _lineEx = 1;
            _lineEy = 0;
            _linePrevCharX = 0;
            _arrayGlyphOffsets.splice(0, _arrayGlyphOffsets.length);

            _lineX = stream.GetDouble();
            _lineY = stream.GetDouble();
            if (stream.GetChar())
            {
                _lineEx = stream.GetDouble();
                _lineEy = stream.GetDouble();
            }
            _lineAscent  = stream.GetDouble();
            _lineDescent = stream.GetDouble();
            _lineWidth   = stream.GetDouble();

            let nChars = stream.GetLong();
            for (let i = 0; i < nChars; ++i)
            {
                if (i)
                    _linePrevCharX += stream.GetDouble();
                _arrayGlyphOffsets[i] = _linePrevCharX;
                stream.Skip(8);
            }

            let startChar = pageIndex == Page1 && iLine == Line1 ? Glyph1 : -2;
            let endChar   = pageIndex == Page2 && iLine == Line2 ? Glyph2 : -1;

            let off1 = 0;
            let off2 = 0;

            if (startChar == -2)
                off1 = 0;
            else if (startChar == -1)
                off1 = _lineWidth;
            else
                off1 = _arrayGlyphOffsets[startChar];

            if (endChar == -2)
                off2 = 0;
            else if (endChar == -1)
                off2 = _lineWidth;
            else
                off2 = _arrayGlyphOffsets[endChar];

            if (off2 <= off1)
                continue;

            // в принципе код один и тот же. Но почти всегда линии горизонтальные.
            // а для горизонтальной линии все можно пооптимизировать
            if (_lineEx == 1 && _lineEy == 0)
            {
                let _x = (x + dKoefX * (_lineX + off1));
                let _r = (x + dKoefX * (_lineX + off2));
                let _y = (y + dKoefY * (_lineY - _lineAscent));
                let _b = (y + dKoefY * (_lineY + _lineDescent));

                overlay.CheckPoint(_x, _y);
                overlay.CheckPoint(_r, _b);

                overlay.m_oContext.rect(_x,_y,_r-_x,_b-_y);
            }
            else
            {
                // определяем точки descent линии
                let ortX = -_lineEy;
                let ortY = _lineEx;

                let _dx = _lineX + ortX * _lineDescent;
                let _dy = _lineY + ortY * _lineDescent;

                let _x1 = _dx + off1 * _lineEx;
                let _y1 = _dy + off1 * _lineEy;

                let _x2 = _x1 - ortX * (_lineAscent + _lineDescent);
                let _y2 = _y1 - ortY * (_lineAscent + _lineDescent);

                let _x3 = _x2 + (off2 - off1) * _lineEx;
                let _y3 = _y2 + (off2 - off1) * _lineEy;

                let _x4 = _x3 + ortX * (_lineAscent + _lineDescent);
                let _y4 = _y3 + ortY * (_lineAscent + _lineDescent);

                _x1 = (x + dKoefX * _x1);
                _x2 = (x + dKoefX * _x2);
                _x3 = (x + dKoefX * _x3);
                _x4 = (x + dKoefX * _x4);

                _y1 = (y + dKoefY * _y1);
                _y2 = (y + dKoefY * _y2);
                _y3 = (y + dKoefY * _y3);
                _y4 = (y + dKoefY * _y4);

                overlay.CheckPoint(_x1, _y1);
                overlay.CheckPoint(_x2, _y2);
                overlay.CheckPoint(_x3, _y3);
                overlay.CheckPoint(_x4, _y4);

                let ctx = overlay.m_oContext;
                ctx.moveTo(_x1, _y1);
                ctx.lineTo(_x2, _y2);
                ctx.lineTo(_x3, _y3);
                ctx.lineTo(_x4, _y4);
                ctx.closePath();
            }
        }
    };
    CFile.prototype.copySelection = function(pageIndex, _text_format)
    {
        let stream = this.getPageTextStream(pageIndex);
        if (!stream || !this.isSelectionUse())
            return "";

        let selection = this.sortSelection();
        let Page1 = selection.Page1;
        let Page2 = selection.Page2;
        let Line1 = selection.Line1;
        let Line2 = selection.Line2;
        let Glyph1 = selection.Glyph1;
        let Glyph2 = selection.Glyph2;

        if (Page1 > pageIndex || Page2 < pageIndex)
            return "";

        let ret = "";
        let iLine = -1;

        let startLine = pageIndex == Page1 ? Line1 : 0;
        let endLine   = pageIndex == Page2 ? Line2 : Infinity;

        while (stream.pos < stream.size)
        {
            iLine++;
            if (iLine < startLine)
            {
                stream.Skip(8);
                if (stream.GetChar())
                    stream.Skip(8);
                stream.Skip(12);
                stream.Skip(12 * stream.GetLong() - 4);
                continue;
            }
            if (iLine > endLine)
                break;

            let startChar = pageIndex == Page1 && iLine == Line1 ? Glyph1 : 0;
            let endChar   = pageIndex == Page2 && iLine == Line2 ? Glyph2 : Infinity;

            if (startChar == -2)
                startChar = 0;
            else if (startChar == -1)
                startChar = Infinity;

            if (endChar == -2)
                endChar = 0;
            else if (endChar == -1)
                endChar = Infinity;

            if (endChar <= startChar)
                continue;

            let textLine = "<p><span>";

            stream.Skip(8);
            if (stream.GetChar())
                stream.Skip(8);
            stream.Skip(12);

            let nChars = stream.GetLong();
            for (let i = 0; i < nChars; ++i)
            {
                if (i)
                    stream.Skip(4);
                let nChar = stream.GetLong();
                stream.Skip(4);

                if (i < startChar || i >= endChar)
                    continue;

                let _char = nChar == 0xFFFF ? ' ' : String.fromCodePoint(nChar);
                textLine += _char;

                if (_text_format)
                    _text_format.Text += _char;
            }

            textLine += "</span></p>";

            if (_text_format)
                _text_format.Text += "\n";

            ret += textLine;
        }
        return ret;
    };
    CFile.prototype.copy = function(_text_format)
    {
        let sel = this.Selection;
        let page1 = sel.Page1;
        let page2 = sel.Page2;
        if (page2 < page1)
        {
            page1 = page2;
            page2 = sel.Page1;
        }

        let ret = "<div>";
        for (let i = page1; i <= page2; ++i)
        {
            if (this.pages[i].isRecognized)
                continue;
            ret += this.copySelection(i, _text_format);
        }
        ret += "</div>";

        return ret;
    };
    CFile.prototype.onUpdateOverlay = function()
    {
        this.viewer.onUpdateOverlay();
    };
    CFile.prototype.onUpdateSelection = function()
    {
        if (this.viewer.Api)
            this.viewer.Api.sendEvent("asc_onSelectionEnd");
    };

    // SEARCH
    CFile.prototype.startSearch = function(text)
    {
        this.viewer.StartSearch();

        this.SearchInfo.Text = text;
        this.SearchInfo.Page = 0;

        var oThis = this;
        this.SearchInfo.Id = setTimeout(function(){oThis.onSearchPage();}, 1);
    };
    CFile.prototype.onSearchPage = function()
    {
        this.SearchPage(this.SearchInfo.Page, this.SearchInfo.Text);
        this.SearchInfo.Page++;

        if (this.SearchInfo.Page >= this.pages.length)
        {
            this.stopSearch();
            return;
        }

        var oThis = this;
        this.SearchInfo.Id = setTimeout(function(){oThis.onSearchPage();}, 1);
    };
    CFile.prototype.stopSearch = function()
    {
        if (null != this.SearchInfo.Id)
        {
            clearTimeout(this.SearchInfo.Id);
            this.SearchInfo.Id = null;
        }
        this.viewer.EndSearch(false);
    };

    // класс элемента совпадения при поиске на странице
    function PdfPageMatch() {
        Array.apply(null, arguments);
        
        this.pdfPageMatch = true;
    }
    PdfPageMatch.prototype = Object.create(Array.prototype);
    PdfPageMatch.prototype.constructor = PdfPageMatch;

    PdfPageMatch.prototype.Get_AbsolutePage = function() {
        if (this[0])
            return this[0].PageNum;
        return -1;
    };
    PdfPageMatch.prototype.GetTextAroundSearchResult = function(nId) {
        let oSearchEngine = Asc.editor.getPDFDoc().SearchEngine;
        let aMatches = oSearchEngine.Elements[nId];
        let oPart, oLine;
        let aResult = ["", "", ""];
        // найденный текст может быть разбит на части (строки)
        for (let nPart = 0; nPart < aMatches.length; nPart++) {
            oPart = aMatches[nPart];
            // знаем в какой строке было найдено совпадение
            oLine = oSearchEngine.PagesLines[oPart.PageNum][oPart.LineNum];

            if (nPart == 0 && aMatches.length == 1) {
                aResult[0] = oLine.slice(0, oPart.Char1);
                aResult[1] = oPart.Text;
                aResult[2] = oLine.slice(oPart.Char2);
            }
            else if (nPart == 0) {
                aResult[0] = oLine.slice(0, oPart.Char1);
                aResult[1] = oPart.Text;
            }
            else if (nPart == aMatches.length - 1) {
                aResult[1] += oPart.Text;
                aResult[2] = oLine.slice(oPart.Char2);
            }
            else {
                aResult[1] += oPart.Text;
            }
        }

        return aResult;
    };

    CFile.prototype.searchPage = function(pageIndex)
    {
        let oSearchEngine = Asc.editor.getPDFDoc().SearchEngine;
        let oResult = {
            matches:    [],
            pageLines:  []
        };

        let searchText = oSearchEngine.Text;
        if (0 == searchText.length)
            return oResult;

        let stream = this.getPageTextStream(pageIndex);
        if (!stream)
            return oResult;
        
        if (!oSearchEngine.MatchCase)
            searchText = searchText.toLowerCase();

        // textline parameters
        let _lineEx = 1;
        let _lineEy = 0;
        let _lineAscent  = 0;
        let _lineDescent = 0;
        
        let curLine = "";

        while (stream.pos < stream.size)
        {
            _lineEx = 1;
            _lineEy = 0;
            curLine = "";

            //curLine.X = stream.GetDouble();
            //curLine.Y = stream.GetDouble();
            stream.Skip(8);
            if (stream.GetChar())
            {
                stream.Skip(8);
                //_lineEx = stream.GetDouble();
                //_lineEy = stream.GetDouble();
                //curLine.Ex = _lineEx;
                //curLine.Ey = _lineEy;
            }
            stream.Skip(12);
            //_lineAscent  = stream.GetDouble();
            //_lineDescent = stream.GetDouble();
            //curLine.X += _lineAscent * _lineEy;
            //curLine.Y -= _lineAscent * _lineEx;
            //curLine.H = _lineAscent + _lineDescent;
            //curLine.W = stream.GetDouble();

            let nChars = stream.GetLong();
            for (let i = 0; i < nChars; ++i)
            {
                if (i)
                    stream.Skip(4);
                let nChar = stream.GetLong();
                curLine += (nChar == 0xFFFF ? " " : String.fromCharCode(nChar));
                stream.Skip(4);
            }

            oResult.pageLines[oResult.pageLines.length] = curLine;
        }

        stream.pos = 0;
        let oMatch = {};
        let _lineX = 0;
        let _lineY = 0;
        let _numLine = 0;
        let _linePos = 0;
        let _startChar = 0;
        let _predChar = 0;
        let _lineWidth = 0;
        let _linePrevCharX = 0;
        let _ignoreFirstSpace = true;
        let _arrayGlyphOffsets = [];
        let _skip = false;

        let posInText = 0;
        if (searchText[posInText] == ' ')
        {
            for (let i = posInText; i < searchText.length; ++i)
            {
                if (searchText[i] == ' ')
                    posInText++;
                else
                    break;
            }
        }
        let PosStartText = posInText;

        while (stream.pos < stream.size)
        {
            if (searchText[posInText] == ' ')
            {
                for (let i = posInText; i < searchText.length; ++i)
                {
                    if (searchText[i] == ' ')
                        posInText++;
                    else
                        break;
                }
            }

            _ignoreFirstSpace = true;
            _linePos = stream.pos;
            stream.Skip(8);
            if (stream.GetChar())
                stream.Skip(8);
            stream.Skip(12);

            let nChars = stream.GetLong();
            for (let i = 0; i < nChars; ++i)
            {
                if (i)
                    stream.Skip(4);
                let nChar = stream.GetLong();
                stream.Skip(4);

                if (i < _startChar || (_ignoreFirstSpace && (nChar == 0xFFFF || nChar == 32)))
                {
                    _predChar = nChar;
                    continue;
                }
                _ignoreFirstSpace = false;

                if (nChar == 0xFFFF)
                    nChar = 32;

                let cChar = String.fromCodePoint(nChar);
                if (!oSearchEngine.MatchCase)
                    cChar = cChar.toLowerCase();

                if (searchText[posInText] != cChar)
                {
                    if (oMatch.Line != undefined)
                    {
                        _skip = true;
                        break;
                    }
                    _predChar = nChar;
                    continue;
                }

                if (posInText == PosStartText)
                { // Начало совпадения
                    oMatch.Line = _numLine;
                    oMatch.Char = i;
                    oMatch.StreamPos = _linePos;
                    oMatch.StartWhole = (i == 0 || _predChar == 0xFFFF || _predChar == 32 || undefined != AscCommon.g_aPunctuation[_predChar]);
                }
                _predChar = nChar;

                if (++posInText == searchText.length)
                { // Полное совпадение
                    if (oSearchEngine.Word)
                    {
                        _skip = true;
                        if (oMatch.StartWhole)
                        {
                            if (i + 1 == nChars)
                                _skip = false;
                            else
                            {
                                stream.Skip(4);
                                nChar = stream.GetLong();
                                stream.Skip(4);
    
                                if (nChar == 0xFFFF || nChar == 32 || undefined != AscCommon.g_aPunctuation[nChar])
                                    _skip = false;
                            }
                        }
                        if (_skip)
                            break;
                    }
                    let rects = new PdfPageMatch();
                    // Добавление всех областей совпадения от oMatch до текущего
                    let _endChar = i + 1;
                    if (_endChar == nChars)
                        _endChar = -1;
                    GetMatches(rects, stream, oMatch, _numLine, _endChar, stream.pos);
                    oResult.matches.push(rects);
                    _skip = true;
                    break;
                }
            }
            _startChar = 0;
            _numLine++;
            _predChar = 0;
            if (_skip)
            { // Возвращаемся к началу совпадения
                _numLine = oMatch.Line;
                _startChar = oMatch.Char + 1;
                stream.pos = oMatch.StreamPos;
                oMatch = {};
                posInText = PosStartText;
                _skip = false;
            }
        }

        function GetMatches(rects, stream, oMatch, curLine, curChar, curStreamPos)
        {
            _numLine = oMatch.Line;
            stream.pos = oMatch.StreamPos;
            let _text = "";
            while (stream.pos < curStreamPos)
            {
                let startChar = _numLine == oMatch.Line ? oMatch.Char : -2;
                let endChar   = _numLine == curLine     ? curChar     : -1;

                if (startChar == -2)
                    startChar = 0;
                else if (startChar == -1)
                    startChar = Infinity;

                if (endChar == -2)
                    endChar = 0;
                else if (endChar == -1)
                    endChar = Infinity;

                _text = "";
                _lineEx = 1;
                _lineEy = 0;
                _linePrevCharX = 0;
                _arrayGlyphOffsets.splice(0, _arrayGlyphOffsets.length);

                _lineX = stream.GetDouble();
                _lineY = stream.GetDouble();
                if (stream.GetChar())
                {
                    _lineEx = stream.GetDouble();
                    _lineEy = stream.GetDouble();
                }
                _lineAscent  = stream.GetDouble();
                _lineDescent = stream.GetDouble();
                _lineWidth   = stream.GetDouble();
                let nChars = stream.GetLong();
                for (let i = 0; i < nChars; ++i)
                {
                    if (i)
                        _linePrevCharX += stream.GetDouble();
                    _arrayGlyphOffsets[i] = _linePrevCharX;
                    let nChar = stream.GetLong();
                    stream.Skip(4);
                    if (i >= endChar)
                        break;
                    if (i >= startChar)
                        _text += (nChar == 0xFFFF ? " " : String.fromCharCode(nChar));
                }

                let off1 = _arrayGlyphOffsets[startChar];
                let off2 = _arrayGlyphOffsets[endChar];
                if (startChar == Infinity)
                    off1 = _lineWidth;
                if (endChar == Infinity)
                    off2 = _lineWidth;

                if (off2 <= off1)
                    continue;

                rects.push({
                    PageNum : pageIndex,
                    LineNum: _numLine,
                    Char1: startChar,
                    Char2: endChar,
                    X : _lineX + _lineAscent * _lineEy + off1 * _lineEx,
                    Y : _lineY - _lineAscent * _lineEx + off1 * _lineEy,
                    W : off2 - off1,
                    H : _lineAscent + _lineDescent,
                    Ex : _lineEx,
                    Ey : _lineEy,
                    Text: _text
                });

                _numLine++;
            }
        }

        return oResult;
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
            file.type = file.nativeFile["getType"]();

            file.nativeFile["onRepaintPages"] = function(pages) {
                file.onRepaintPages && file.onRepaintPages(pages);
            };
            file.nativeFile["onRepaintAnnotations"] = function(pages) {
                file.onRepaintAnnotations && file.onRepaintAnnotations(pages);
            };
            file.nativeFile["onRepaintForms"] = function(pages) {
                file.onRepaintForms && file.onRepaintForms(pages);
            };

            file.nativeFile["onUpdateStatistics"] = function(par, word, symbol, space) {
                file.onUpdateStatistics && file.onUpdateStatistics(par, word, symbol, space);
            };
            file.nativeFile["isPunctuation"] = function(unicode) {
                return AscCommon.g_aPunctuation[unicode];
            };
            file.pages = file.nativeFile["getPages"]();

            for (var i = 0, len = file.pages.length; i < len; i++)
            {
                var page = file.pages[i];
                
                page.W              = page["W"];
                page.H              = page["H"];
                page.Dpi            = page["Dpi"];
                page.originIndex    = page["originIndex"]; // исходный индекс в файле
                page.originRotate   = page["Rotate"];
                page.Rotate         = page["Rotate"];
            }
            file.originalPagesCount = file.pages.length;

            //file.cacheManager = new AscCommon.CCacheManager();
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
            file.type = file.nativeFile["getType"]();

            file.nativeFile["onRepaintPages"] = function(pages) {
                file.onRepaintPages && file.onRepaintPages(pages);
            };
            file.nativeFile["onUpdateStatistics"] = function(par, word, symbol, space) {
                file.onUpdateStatistics && file.onUpdateStatistics(par, word, symbol, space);
            };
            file.nativeFile["isPunctuation"] = function(unicode) {
                return AscCommon.g_aPunctuation[unicode];
            };
            file.pages = file.nativeFile["getPages"]();

            for (var i = 0, len = file.pages.length; i < len; i++)
            {
                var page = file.pages[i];
                page.W              = page["W"];
                page.H              = page["H"];
                page.Dpi            = page["Dpi"];
                page.originIndex    = page["originIndex"]; // исходный индекс в файле
                page.originRotate   = page["Rotate"];
                page.Rotate         = page["Rotate"];
            }
            file.originalPagesCount = file.pages.length;
            
            //file.cacheManager = new AscCommon.CCacheManager();
        }
    };
	window["AscViewer"].createEmptyFile = function()
	{
		return new CFile();
	};

    //--------------------------------------------------------export----------------------------------------------------
	window['AscPDF'].PdfPageMatch = PdfPageMatch;

})(window, undefined);
