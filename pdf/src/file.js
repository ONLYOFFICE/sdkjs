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
        this.GetUChar = function()
        {
            if (this.pos >= this.size)
                return 0;
            return this.data[this.pos++];
        };
        this.GetChar = function()
        {
            if (this.pos >= this.size)
                return 0;
            var m = this.data[this.pos++];
            if (m > 0x7F)
                m -= 256;
            return m;
        };

        // 2 byte
        this.GetUShort = function()
        {
            if (this.pos + 1 >= this.size)
                return 0;
            return (this.data[this.pos++] | this.data[this.pos++] << 8);
        };
        this.GetShort = function()
        {
            if (this.pos + 1 >= this.size)
                return 0;
            var _c = (this.data[this.pos++] | this.data[this.pos++] << 8);

            if (_c > 0x7FFF)
                return _c - 65536;
            return _c;
        };

        // 4 byte
        this.GetULong = function()
        {
            if (this.pos + 3 >= this.size)
                return 0;
            var s = (this.data[this.pos++] | this.data[this.pos++] << 8 | this.data[this.pos++] << 16 | this.data[this.pos++] << 24);
            if (s < 0)
                s += (0xFFFFFFFF + 1);
            return s;
        };
        this.GetLong = function()
        {
            return (this.data[this.pos++] | this.data[this.pos++] << 8 | this.data[this.pos++] << 16 | this.data[this.pos++] << 24);
        };

        // double
        this.GetDouble = function()
        {
            return this.GetLong() / 10000;
        };
        this.GetDouble2 = function()
        {
            return this.GetShort() / 100;
        };
    }

    function CSpan()
    {
        this.fontName = 0;
        this.fontSize = 0;

        this.colorR = 0;
        this.colorG = 0;
        this.colorB = 0;

        this.inner = "";

        this.CreateDublicate = function()
        {
            var ret = new CSpan();

            ret.fontName = this.fontName;
            ret.fontSize = this.fontSize;

            ret.colorR = this.colorR;
            ret.colorG = this.colorG;
            ret.colorB = this.colorB;

            ret.inner = this.inner;

            return ret;
        }
    }

    var supportImageDataConstructor = (AscCommon.AscBrowser.isIE && !AscCommon.AscBrowser.isIeEdge) ? false : true;

    function CFile()
    {
    	this.nativeFile = 0;
    	this.pages = [];
    	this.zoom = 1;
    	this.isUse3d = false;
    	this.cacheManager = null;
    	this.logging = true;

    	this.Selection = {
            Page1 : 0,
            Line1 : 0,
            Glyph1 : 0,

            Page2 : 0,
            Line2 : 0,
            Glyph2 : 0,

            IsSelection : false
        };
        this.SearchResults = {
            IsSearch    : false,
            Text        : "",
            MachingCase : false,
            Pages       : [],
            CurrentPage : -1,
            Current     : -1,
            Show        : false,
            Count       : 0
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
    CFile.prototype.getDocumentInfo = function()
    {
        return this.nativeFile ? this.nativeFile["getDocumentInfo"]() : null;
    };

    CFile.prototype.getPage = function(pageIndex, width, height, isNoUseCacheManager, backgroundColor)
    {
        if (!this.nativeFile)
            return null;
        if (pageIndex < 0 || pageIndex >= this.pages.length)
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

    // TEXT
    CFile.prototype.logTextCommands = function(commands)
    {
        var stream = new TextStreamReader(commands, commands.length);
        var lineCharCount = 0;
        var lineGidExist = false;
        var lineText = "";
        while (stream.pos < stream.size)
        {
            var command = stream.GetUChar();

            switch (command)
            {
                case 41: // ctFontName
                {
                    stream.Skip(12);
                    break;
                }
                case 22: // ctBrushColor1
                {
                    stream.Skip(4);
                    break;
                }
                case 80: // ctDrawText
                {
                    if (0 != lineCharCount)
                        stream.Skip(2);

                    lineCharCount++;

                    var char = stream.GetUShort();
                    if (char !== 0xFFFF)
                        lineText += String.fromCharCode(char);
                    if (lineGidExist)
                        stream.Skip(2);

                    stream.Skip(2);
                    break;
                }
                case 160: // ctCommandTextLine
                {
                    lineText = "";
                    lineCharCount = 0;
                    var mask = stream.GetUChar();
                    stream.Skip(8);

                    if ((mask & 0x01) == 0)
                    {
                        stream.Skip(8);
                    }

                    stream.Skip(8);

                    if ((mask & 0x04) != 0)
                        stream.Skip(4);

                    if ((mask & 0x02) != 0)
                        lineGidExist = true;
                    else
                        lineGidExist = false;

                    break;
                }
                case 161: // ctCommandTextTransform
                {
                    // text transform
                    stream.Skip(16);
                    break;
                }
                case 162: // ctCommandTextLineEnd
                {
                    console.log(lineText);
                    break;
                }
                default:
                {
                    stream.pos = stream.size;
                }
            }
        }
    };

    CFile.prototype.onMouseDown = function(pageIndex, x, y)
    {
        var ret = this.getNearestPos(pageIndex, x, y);
        var sel = this.Selection;

        sel.Page1  = pageIndex;
        sel.Line1  = ret.Line;
        sel.Glyph1 = ret.Glyph;

        sel.Page2  = pageIndex;
        sel.Line2  = ret.Line;
        sel.Glyph2 = ret.Glyph;

        sel.IsSelection = true;

        this.onUpdateSelection();
        this.onUpdateOverlay();
    };

    CFile.prototype.onMouseMove = function(pageIndex, x, y)
    {
        if (false === this.Selection.IsSelection)
            return;

        var ret = this.getNearestPos(pageIndex, x, y);
        var sel = this.Selection;

        sel.Page2  = pageIndex;
        sel.Line2  = ret.Line;
        sel.Glyph2 = ret.Glyph;

        this.onUpdateOverlay();
    };

    CFile.prototype.onMouseUp = function()
    {
        this.Selection.IsSelection = false;
        this.onUpdateSelection();
        this.onUpdateOverlay();
    };

    CFile.prototype.getPageTextStream = function(pageIndex)
    {
        var textCommands = this.pages[pageIndex].text;
        if (!textCommands)
            return null;

        return new TextStreamReader(textCommands, textCommands.length);
    };

    CFile.prototype.getNearestPos = function(pageIndex, x, y)
    {
        var stream = this.getPageTextStream(pageIndex);
        if (!stream)
            return { Line : -1, Glyph : -1 };

        // textline parameters
        var _line = -1;
        var _glyph = -1;
        var _minDist = 0xFFFFFF;

        // textline parameters
        var _lineX = 0;
        var _lineY = 0;
        var _lineEx = 0;
        var _lineEy = 0;
        var _lineAscent = 0;
        var _lineDescent = 0;
        var _lineWidth = 0;
        var _lineGidExist = false;
        var _linePrevCharX = 0;
        var _lineCharCount = 0;
        var _lineLastGlyphWidth = 0;
        var _arrayGlyphOffsets = [];

        var _numLine = -1;

        var _lenGls = 0;
        var tmp = 0;

        while (stream.pos < stream.size)
        {
            var command = stream.GetUChar();

            switch (command)
            {
                case 41:  // ctFontName
                {
                    stream.Skip(12);
                    break;
                }
                case 22: // ctBrushColor1
                {
                    stream.Skip(4);
                    break;
                }
                case 80: // ctDrawText
                {
                    if (0 != _lineCharCount)
                        _linePrevCharX += stream.GetDouble2();

                    _arrayGlyphOffsets[_lineCharCount] = _linePrevCharX;

                    _lineCharCount++;

                    if (_lineGidExist)
                        stream.Skip(4);
                    else
                        stream.Skip(2);

                    if (0 == _lineWidth)
                        _lineLastGlyphWidth = stream.GetDouble2();
                    else
                        stream.Skip(2);

                    break;
                }
                case 160:
                {
                    // textline
                    _linePrevCharX = 0;
                    _lineCharCount = 0;
                    _lineWidth = 0;

                    _arrayGlyphOffsets.splice(0, _arrayGlyphOffsets.length);

                    ++_numLine;

                    var mask = stream.GetUChar();
                    _lineX = stream.GetDouble();
                    _lineY = stream.GetDouble();

                    if ((mask & 0x01) != 0)
                    {
                        _lineEx = 1;
                        _lineEy = 0;
                    }
                    else
                    {
                        _lineEx = stream.GetDouble();
                        _lineEy = stream.GetDouble();
                    }

                    _lineAscent = stream.GetDouble();
                    _lineDescent = stream.GetDouble();

                    if ((mask & 0x04) != 0)
                        _lineWidth = stream.GetDouble();

                    if ((mask & 0x02) != 0)
                        _lineGidExist = true;
                    else
                        _lineGidExist = false;

                    break;
                }
                case 162:
                {
                    // textline end

                    // все подсчитано
                    if (0 == _lineWidth)
                        _lineWidth = _linePrevCharX + _lineLastGlyphWidth;

                    // в принципе код один и тот же. Но почти всегда линии горизонтальные.
                    // а для горизонтальной линии все можно пооптимизировать
                    if (_lineEx == 1 && _lineEy == 0)
                    {
                        var _distX = x - _lineX;
                        if (y >= (_lineY - _lineAscent) && y <= (_lineY + _lineDescent) && _distX >= 0 && _distX <= _lineWidth)
                        {
                            // попали внутрь линии. Теперь нужно найти глиф
                            _line = _numLine;

                            _lenGls = _arrayGlyphOffsets.length;
                            for (_glyph = 0; _glyph < _lenGls; _glyph++)
                            {
                                if (_arrayGlyphOffsets[_glyph] > _distX)
                                    break;
                            }
                            if (_glyph > 0)
                                --_glyph;

                            return { Line : _line, Glyph : _glyph };
                        }

                        if (_distX >= 0 && _distX <= _lineWidth)
                            tmp = Math.abs(y - _lineY);
                        else if (_distX < 0)
                        {
                            tmp = Math.sqrt((x - _lineX) * (x - _lineX) + (y - _lineY) * (y - _lineY));
                        }
                        else
                        {
                            var _xx1 = _lineX + _lineWidth;
                            tmp = Math.sqrt((x - _xx1) * (x - _xx1) + (y - _lineY) * (y - _lineY));
                        }

                        if (tmp < _minDist)
                        {
                            _minDist = tmp;
                            _line = _numLine;

                            if (_distX < 0)
                                _glyph = -2;
                            else if (_distX > _lineWidth)
                            {
                                _glyph = -1;
                            }
                            else
                            {
                                _lenGls = _arrayGlyphOffsets.length;
                                for (_glyph = 0; _glyph < _lenGls; _glyph++)
                                {
                                    if (_arrayGlyphOffsets[_glyph] > _distX)
                                        break;
                                }

                                if (_glyph > 0)
                                    _glyph--;
                            }
                        }

                        // Ничего не надо делать, уже найдена более "ближняя" линия
                    }
                    else
                    {
                        // определяем точки descent линии
                        var ortX = -_lineEy;
                        var ortY = _lineEx;

                        var _dx = _lineX + ortX * _lineDescent;
                        var _dy = _lineY + ortY * _lineDescent;

                        // теперь проекции (со знаком) на линию descent
                        var h = -((x - _dx) * ortX + (y - _dy) * ortY);
                        var w = (x - _dx) * _lineEx + (y - _dy) * _lineEy;

                        if (w >= 0 && w <= _lineWidth && h >= 0 && h <= (_lineDescent + _lineAscent))
                        {
                            // попали внутрь линии. Теперь нужно найти глиф
                            _line = _numLine;

                            _lenGls = _arrayGlyphOffsets.length;
                            for (_glyph = 0; _glyph < _lenGls; _glyph++)
                            {
                                if (_arrayGlyphOffsets[_glyph] > w)
                                    break;
                            }

                            if (_glyph > 0)
                                _glyph--;

                            return { Line : _line, Glyph : _glyph };
                        }

                        if (w >= 0 && w <= _lineWidth)
                            tmp = Math.abs(h - _lineDescent);
                        else if (w < 0)
                        {
                            tmp = Math.sqrt((x - _lineX) * (x - _lineX) + (y - _lineY) * (y - _lineY));
                        }
                        else
                        {
                            var _tmpX = _lineX + _lineWidth * _lineEx;
                            var _tmpY = _lineY + _lineWidth * _lineEy;
                            tmp = Math.sqrt((x - _tmpX) * (x - _tmpX) + (y - _tmpY) * (y - _tmpY));
                        }

                        //tmp = Math.abs(h - _lineDescent);
                        if (tmp < _minDist)
                        {
                            _minDist = tmp;
                            _line = _numLine;

                            if (w < 0)
                                _glyph = -2;
                            else if (w > _lineWidth)
                                _glyph = -1;
                            else
                            {
                                _lenGls = _arrayGlyphOffsets.length;
                                for (_glyph = 0; _glyph < _lenGls; _glyph++)
                                {
                                    if (_arrayGlyphOffsets[_glyph] > w)
                                        break;
                                }

                                if (_glyph > 0)
                                    _glyph--;
                            }
                        }

                        // Ничего не надо делать, уже найдена более "ближняя" линия
                    }

                    break;
                }
                case 161:
                {
                    // text transform
                    stream.Skip(16);
                    break;
                }
                default:
                {
                    stream.pos = stream.size;
                }
            }
        }

        return { Line : _line, Glyph : _glyph };
    };

    CFile.prototype.drawSelection = function(pageIndex, overlay, x, y, width, height)
    {
        var stream = this.getPageTextStream(pageIndex);
        if (!stream)
            return;

        var sel = this.Selection;
        var Page1 = 0;
        var Page2 = 0;
        var Line1 = 0;
        var Line2 = 0;
        var Glyph1 = 0;
        var Glyph2 = 0;

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

        if (Page1 > pageIndex || Page2 < pageIndex)
            return;

        if (Page1 < pageIndex)
        {
            Page1 = pageIndex;
            Line1 = 0;
            Glyph1 = -2;
        }
        var bIsFillToEnd = false;
        if (Page2 > pageIndex)
            bIsFillToEnd = true;

        // textline parameters
        var _lineX = 0;
        var _lineY = 0;
        var _lineEx = 0;
        var _lineEy = 0;
        var _lineAscent = 0;
        var _lineDescent = 0;
        var _lineWidth = 0;
        var _lineGidExist = false;
        var _linePrevCharX = 0;
        var _lineCharCount = 0;
        var _lineLastGlyphWidth = 0;
        var _arrayGlyphOffsets = [];

        var _numLine = -1;

        var dKoefX = width / this.pages[pageIndex].W;
        var dKoefY = height / this.pages[pageIndex].H;
        dKoefX *= (this.pages[pageIndex].Dpi / 25.4);
        dKoefY *= (this.pages[pageIndex].Dpi / 25.4);

        while (stream.pos < stream.size)
        {
            var command = stream.GetUChar();

            switch (command)
            {
                case 41:
                {
                    stream.Skip(12);
                    break;
                }
                case 22:
                {
                    stream.Skip(4);
                    break;
                }
                case 80:
                {
                    if (0 != _lineCharCount)
                        _linePrevCharX += stream.GetDouble2();

                    _arrayGlyphOffsets[_lineCharCount] = _linePrevCharX;

                    _lineCharCount++;

                    if (_lineGidExist)
                        stream.Skip(4);
                    else
                        stream.Skip(2);

                    if (0 == _lineWidth)
                        _lineLastGlyphWidth = stream.GetDouble2();
                    else
                        stream.Skip(2);

                    break;
                }
                case 160:
                {
                    // textline
                    _linePrevCharX = 0;
                    _lineCharCount = 0;
                    _lineWidth = 0;

                    _arrayGlyphOffsets.splice(0, _arrayGlyphOffsets.length);

                    ++_numLine;

                    var mask = stream.GetUChar();
                    _lineX = stream.GetDouble();
                    _lineY = stream.GetDouble();

                    if ((mask & 0x01) != 0)
                    {
                        _lineEx = 1;
                        _lineEy = 0;
                    }
                    else
                    {
                        _lineEx = stream.GetDouble();
                        _lineEy = stream.GetDouble();
                    }

                    _lineAscent = stream.GetDouble();
                    _lineDescent = stream.GetDouble();

                    if ((mask & 0x04) != 0)
                        _lineWidth = stream.GetDouble();

                    if ((mask & 0x02) != 0)
                        _lineGidExist = true;
                    else
                        _lineGidExist = false;

                    break;
                }
                case 162:
                {
                    // textline end
                    var off1 = 0;
                    var off2 = 0;

                    if (_numLine < Line1)
                        break;
                    if (_numLine > Line2 && !bIsFillToEnd)
                        return;

                    // все подсчитано
                    if (0 == _lineWidth)
                        _lineWidth = _linePrevCharX + _lineLastGlyphWidth;

                    if (Line1 == _numLine)
                    {
                        if (-2 == Glyph1)
                            off1 = 0;
                        else if (-1 == Glyph1)
                            off1 = _lineWidth;
                        else
                            off1 = _arrayGlyphOffsets[Glyph1];
                    }
                    if (bIsFillToEnd || Line2 != _numLine)
                        off2 = _lineWidth;
                    else
                    {
                        if (Glyph2 == -2)
                            off2 = 0;
                        else if (Glyph2 == -1)
                            off2 = _lineWidth;
                        else
                        {
                            off2 = _arrayGlyphOffsets[Glyph2];
                            /*
                            if (Glyph2 >= (_arrayGlyphOffsets.length - 1))
                                off2 = _lineWidth;
                            else
                                off2 = _arrayGlyphOffsets[Glyph2 + 1];
                            */
                        }
                    }

                    if (off2 <= off1)
                        break;

                    // в принципе код один и тот же. Но почти всегда линии горизонтальные.
                    // а для горизонтальной линии все можно пооптимизировать
                    if (_lineEx == 1 && _lineEy == 0)
                    {
                        var _x = (x + dKoefX * (_lineX + off1)) >> 0;
                        var _r = (x + dKoefX * (_lineX + off2)) >> 0;
                        var _y = (y + dKoefY * (_lineY - _lineAscent)) >> 0;
                        var _b = (y + dKoefY * (_lineY + _lineDescent)) >> 0;

                        if (_x < overlay.min_x)
                            overlay.min_x = _x;
                        if (_r > overlay.max_x)
                            overlay.max_x = _r;

                        if (_y < overlay.min_y)
                            overlay.min_y = _y;
                        if (_b > overlay.max_y)
                            overlay.max_y = _b;

                        overlay.m_oContext.rect(_x,_y,_r-_x,_b-_y);
                    }
                    else
                    {
                        // определяем точки descent линии
                        var ortX = -_lineEy;
                        var ortY = _lineEx;

                        var _dx = _lineX + ortX * _lineDescent;
                        var _dy = _lineY + ortY * _lineDescent;

                        var _x1 = _dx + off1 * _lineEx;
                        var _y1 = _dy + off1 * _lineEy;

                        var _x2 = _x1 - ortX * (_lineAscent + _lineDescent);
                        var _y2 = _y1 - ortY * (_lineAscent + _lineDescent);

                        var _x3 = _x2 + (off2 - off1) * _lineEx;
                        var _y3 = _y2 + (off2 - off1) * _lineEy;

                        var _x4 = _x3 + ortX * (_lineAscent + _lineDescent);
                        var _y4 = _y3 + ortY * (_lineAscent + _lineDescent);

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

                        var ctx = overlay.m_oContext;
                        ctx.moveTo(_x1, _y1);
                        ctx.lineTo(_x2, _y2);
                        ctx.lineTo(_x3, _y3);
                        ctx.lineTo(_x4, _y4);
                        ctx.closePath();
                    }

                    break;
                }
                case 161:
                {
                    // text transform
                    stream.Skip(16);
                    break;
                }
                default:
                {
                    stream.pos = stream.size;
                }
            }
        }
    };

    CFile.prototype.copySelection = function(pageIndex, _text_format)
    {
        var stream = this.getPageTextStream(pageIndex);
        if (!stream)
            return;

        var ret = "";

        var sel = this.Selection;
        var Page1 = 0;
        var Page2 = 0;
        var Line1 = 0;
        var Line2 = 0;
        var Glyph1 = 0;
        var Glyph2 = 0;

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
        else if (sel.Page1 == sel.Page2)
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

                if (((sel.Glyph1 != -1) && (sel.Glyph1 < sel.Glyph2)) || (-1 == sel.Glyph2))
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

        if (Page1 > pageIndex  || Page2 < pageIndex)
            return;

        if (Page1 < pageIndex)
        {
            Page1 = pageIndex;
            Line1 = 0;
            Glyph1 = -2;
        }
        var bIsFillToEnd = false;
        if (Page2 > pageIndex)
            bIsFillToEnd = true;


        var lineSpans = [];
        var curSpan = new CSpan();
        var isChangeSpan = false;

        var _lineCharCount = 0;
        var _lineGidExist = false;

        var _numLine = -1;

        while (stream.pos < stream.size)
        {
            var command = stream.GetUChar();

            switch (command)
            {
                case 41:
                {
                    curSpan.fontName = stream.GetULong();
                    stream.Skip(4);
                    curSpan.fontSize = stream.GetDouble();
                    isChangeSpan = true;
                    break;
                }
                case 22:
                {
                    curSpan.colorR = stream.GetUChar();
                    curSpan.colorG = stream.GetUChar();
                    curSpan.colorB = stream.GetUChar();
                    stream.Skip(1);
                    isChangeSpan = true;
                    break;
                }
                case 80:
                {
                    if (0 != _lineCharCount)
                        stream.Skip(2);

                    _lineCharCount++;
                    if (isChangeSpan)
                    {
                        lineSpans[lineSpans.length] = curSpan.CreateDublicate();
                    }
                    var sp = lineSpans[lineSpans.length - 1];

                    var _char = stream.GetUShort();
                    if (0xFFFF == _char)
                        sp.inner += " ";
                    else
                        sp.inner += String.fromCharCode(_char);

                    if (_lineGidExist)
                        stream.Skip(2);

                    stream.Skip(2);

                    isChangeSpan = false;
                    break;
                }
                case 160:
                {
                    // textline
                    isChangeSpan = true;
                    lineSpans.splice(0, lineSpans.length);
                    _lineCharCount = 0;
                    ++_numLine;

                    var mask = stream.GetUChar();
                    stream.Skip(8);

                    if ((mask & 0x01) == 0)
                    {
                        stream.Skip(8);
                    }

                    stream.Skip(8);

                    if ((mask & 0x04) != 0)
                        stream.Skip(4);

                    if ((mask & 0x02) != 0)
                        _lineGidExist = true;
                    else
                        _lineGidExist = false;

                    break;
                }
                case 162:
                {
                    // textline end
                    // спаны набиты. теперь нужно сформировать линию и сгенерировать нужную строку.
                    if (Line1 <= _numLine && ((!bIsFillToEnd && Line2 >= _numLine) || bIsFillToEnd))
                    {
                        var _g1 = -2;
                        var _g2 = -1;
                        if (Line1 == _numLine)
                        {
                            _g1 = Glyph1;
                        }
                        if (bIsFillToEnd || Line2 != _numLine)
                        {
                            _g2 = -1;
                        }
                        else
                        {
                            _g2 = Glyph2;
                        }

                        if (_g1 != -1 && _g2 != -2)
                        {
                            var textLine = "<p>";

                            if (-2 == _g1 && -1 == _g2)
                            {
                                var countSpans = lineSpans.length;
                                for (var i = 0; i < countSpans; i++)
                                {
                                    textLine += "<span>";
                                    textLine += lineSpans[i].inner;
                                    textLine += "</span>";

                                    if (_text_format)
                                        _text_format.Text += lineSpans[i].inner;
                                }
                            }
                            else
                            {
                                var curIndex = 0;
                                var countSpans = lineSpans.length;
                                for (var i = 0; i < countSpans; i++)
                                {
                                    var old = curIndex;
                                    var start = curIndex;
                                    var end = start + lineSpans[i].inner.length;
                                    curIndex = end;

                                    if (_g1 > start)
                                        start = _g1;
                                    if (_g2 != -1 && _g2 < end)
                                        end = _g2;

                                    if (start > end)
                                        continue;

                                    start -= old;
                                    end -= old;

                                    textLine += "<span>";
                                    textLine += lineSpans[i].inner.substring(start, end);
                                    textLine += "</span>";

                                    if (_text_format)
                                        _text_format.Text += lineSpans[i].inner.substring(start, end);
                                }
                            }

                            textLine += "</p>";

                            if (_text_format)
                                _text_format.Text += "\n";

                            ret += textLine;
                        }
                    }

                    break;
                }
                case 161:
                {
                    // text transform
                    stream.Skip(16);
                    break;
                }
                default:
                {
                    stream.pos = stream.size;
                }
            }
        }
        return ret;
    };

    CFile.prototype.copy = function(_text_format)
    {
        var sel = this.Selection;
        var page1 = sel.Page1;
        var page2 = sel.Page2;

        if (page2 < page1)
        {
            page1 = page2;
            page2 = sel.Page1;
        }

        var ret = "<div>";
        for (var i = page1; i <= page2; i++)
        {
            ret += this.copySelection(i, _text_format);
        }
        ret += "</div>";

        //console.log(ret);
        return ret;
    };

    CFile.prototype.getCountLines = function(pageIndex)
    {
        var stream = this.getPageTextStream(pageIndex);
        if (!stream)
            return -1;

        var _lineGidExist = false;
        var _lineCharCount = 0;
        var _numLine = -1;

        while (stream.pos < stream.size)
        {
            var command = stream.GetUChar();

            switch (command)
            {
                case 41:
                {
                    stream.Skip(12);
                    break;
                }
                case 22:
                {
                    stream.Skip(4);
                    break;
                }
                case 80:
                {
                    if (0 != _lineCharCount)
                        stream.Skip(2);

                    _lineCharCount++;
                    stream.Skip(_lineGidExist ? 6 : 4);
                    break;
                }
                case 160:
                {
                    // textline
                    _lineCharCount = 0;
                    ++_numLine;

                    var mask = stream.GetUChar();
                    stream.Skip(8);

                    if ((mask & 0x01) == 0)
                        stream.Skip(8);

                    stream.Skip(8);

                    if ((mask & 0x04) != 0)
                        stream.Skip(4);

                    if ((mask & 0x02) != 0)
                        _lineGidExist = true;
                    else
                        _lineGidExist = false;

                    break;
                }
                case 162:
                {
                    break;
                }
                case 161:
                {
                    // text transform
                    stream.Skip(16);
                    break;
                }
                default:
                {
                    stream.pos = stream.size;
                }
            }
        }

        return _numLine;
    };

    CFile.prototype.selectAll = function()
    {
        var sel = this.Selection;

        sel.Page1 = 0;
        sel.Line1 = 0;
        sel.Glyph1 = 0;

        sel.Page2 = 0;
        sel.Line2 = 0;
        sel.Glyph2 = 0;

        sel.IsSelection = false;

        var pagesCount = this.pages.length;
        if (0 != pagesCount)
        {
            var lLinesLastPage = this.getCountLines(pagesCount - 1);
            if (1 != pagesCount || 0 != lLinesLastPage)
            {
                sel.Glyph1 = -2;
                sel.Page2 = pagesCount - 1;
                sel.Line2 = lLinesLastPage;
                sel.Glyph2 = -1;
            }
        }

        this.onUpdateSelection();
        this.onUpdateOverlay();
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

    CFile.prototype.searchPage = function(pageIndex)
    {
        var stream = this.getPageTextStream(pageIndex);
        if (!stream)
            return;

        var _searchResults = this.SearchResults;
        var _navRects = _searchResults.Pages[pageIndex];

        var glyphsEqualFound = 0;
        var text = _searchResults.Text;
        var glyphsFindCount = text.length;

        if (!_searchResults.MachingCase)
        {
            text = text.toLowerCase();
        }

        if (0 == glyphsFindCount)
            return;

        var _numLine = -1;
        var _lineGidExist = false;
        var _linePrevCharX = 0;
        var _lineCharCount = 0;
        var _lineLastGlyphWidth = 0;

        var _findLine = 0;
        var _findLineOffsetX = 0;
        var _findLineOffsetR = 0;
        var _findGlyphIndex = 0;

        var _SeekToNextPoint = 0;
        var _SeekLinePrevCharX = 0;

        var arrayLines = [];
        var curLine = null;

        while (stream.pos < stream.size)
        {
            var command = stream.GetUChar();

            switch (command)
            {
                case 41:
                {
                    stream.Skip(12);
                    break;
                }
                case 22:
                {
                    stream.Skip(4);
                    break;
                }
                case 80:
                {
                    if (0 != _lineCharCount)
                        _linePrevCharX += stream.GetDouble2();

                    _lineCharCount++;

                    var _char = stream.GetUShort();
                    if (_lineGidExist)
                        stream.Skip(2);

                    if (0xFFFF == _char)
                        curLine.text += " ";
                    else
                        curLine.text += String.fromCharCode(_char);

                    if (curLine.W != 0)
                        stream.Skip(2);
                    else
                        curLine.W = stream.GetDouble2();

                    break;
                }
                case 160:
                {
                    _linePrevCharX = 0;
                    _lineCharCount = 0;

                    arrayLines[arrayLines.length] = new CLineInfo();
                    curLine = arrayLines[arrayLines.length - 1];

                    var mask = stream.GetUChar();
                    curLine.X = stream.GetDouble();
                    curLine.Y = stream.GetDouble();

                    if ((mask & 0x01) == 1)
                    {
                        var dAscent = stream.GetDouble();
                        var dDescent = stream.GetDouble();

                        curLine.Y -= dAscent;
                        curLine.H = dAscent + dDescent;
                    }
                    else
                    {
                        curLine.Ex = stream.GetDouble();
                        curLine.Ey = stream.GetDouble();

                        var dAscent = stream.GetDouble();
                        var dDescent = stream.GetDouble();

                        curLine.X = curLine.X + dAscent * curLine.Ey;
                        curLine.Y = curLine.Y - dAscent * curLine.Ex;

                        curLine.H = dAscent + dDescent;
                    }

                    if ((mask & 0x04) != 0)
                        curLine.W = stream.GetDouble();

                    if ((mask & 0x02) != 0)
                        _lineGidExist = true;
                    else
                        _lineGidExist = false;

                    break;
                }
                case 162:
                {
                    break;
                }
                case 161:
                {
                    // text transform
                    stream.Skip(16);
                    break;
                }
                default:
                {
                    stream.pos = stream.size;
                }
            }
        }

        // текст заполнен. теперь нужно просто пробегаться и смотреть
        // откуда совпадение началось и где закончилось
        _linePrevCharX = 0;
        _lineCharCount = 0;
        _numLine = 0;

        stream.Seek(0);

        while (stream.pos < stream.size)
        {
            var command = stream.GetUChar();

            switch (command)
            {
                case 41:
                {
                    stream.Skip(12);
                    break;
                }
                case 22:
                {
                    stream.Skip(4);
                    break;
                }
                case 80:
                {
                    if (0 != _lineCharCount)
                        _linePrevCharX += stream.GetDouble2();

                    _lineCharCount++;

                    var _char = stream.GetUShort();
                    if (_lineGidExist)
                        stream.Skip(2);

                    if (0xFFFF == _char)
                        _char = " ".charCodeAt(0);

                    _lineLastGlyphWidth = stream.GetDouble2();

                    var _isFound = false;
                    if (_searchResults.MachingCase)
                    {
                        if (_char == text.charCodeAt(glyphsEqualFound))
                            _isFound = true;
                    }
                    else
                    {
                        var _strMem = String.fromCharCode(_char);
                        _strMem = _strMem.toLowerCase();
                        if (_strMem.charCodeAt(0) == text.charCodeAt(glyphsEqualFound))
                            _isFound = true;
                    }

                    if (_isFound)
                    {
                        if (0 == glyphsEqualFound)
                        {
                            _findLine = _numLine;
                            _findLineOffsetX = _linePrevCharX;
                            _findGlyphIndex = _lineCharCount;

                            _SeekToNextPoint = stream.pos;
                            _SeekLinePrevCharX = _linePrevCharX;
                        }

                        glyphsEqualFound++;
                        _findLineOffsetR = _linePrevCharX + _lineLastGlyphWidth;
                        if (glyphsFindCount == glyphsEqualFound)
                        {
                            var _rects = [];
                            for (var i = _findLine; i <= _numLine; i++)
                            {
                                var ps = 0;
                                if (_findLine == i)
                                    ps = _findLineOffsetX;
                                var pe = arrayLines[i].W;
                                if (i == _numLine)
                                    pe = _findLineOffsetR;

                                var _l = arrayLines[i];
                                if (_l.Ex == 1 && _l.Ey == 0)
                                {
                                    _rects[_rects.length] = { PageNum : pageIndex, X : _l.X + ps, Y : _l.Y, W : pe - ps, H : _l.H };
                                }
                                else
                                {
                                    _rects[_rects.length] = { PageNum : pageIndex, X : _l.X + ps * _l.Ex, Y : _l.Y + ps * _l.Ey, W : pe - ps, H : _l.H, Ex : _l.Ex, Ey : _l.Ey };
                                }
                            }

                            _navRects[_navRects.length] = _rects;

                            // нужно вернуться и попробовать искать со след буквы.
                            glyphsEqualFound = 0;
                            stream.pos = _SeekToNextPoint;
                            _linePrevCharX = _SeekLinePrevCharX;
                            _lineCharCount = _findGlyphIndex;
                            _numLine = _findLine;
                        }
                    }
                    else
                    {
                        if (0 != glyphsEqualFound)
                        {
                            // нужно вернуться и попробовать искать со след буквы.
                            glyphsEqualFound = 0;
                            stream.pos = _SeekToNextPoint;
                            _linePrevCharX = _SeekLinePrevCharX;
                            _lineCharCount = _findGlyphIndex;
                            _numLine = _findLine;
                        }
                    }

                    break;
                }
                case 160:
                {
                    // textline
                    _linePrevCharX = 0;
                    _lineCharCount = 0;

                    var mask = stream.GetUChar();
                    stream.Skip(8);

                    if ((mask & 0x01) == 0)
                        stream.Skip(8);

                    stream.Skip(8);

                    if ((mask & 0x04) != 0)
                        stream.Skip(4);

                    if ((mask & 0x02) != 0)
                        _lineGidExist = true;
                    else
                        _lineGidExist = false;

                    break;
                }
                case 162:
                {
                    ++_numLine;
                    break;
                }
                case 161:
                {
                    // text transform
                    stream.Skip(16);
                    break;
                }
                default:
                {
                    stream.pos = stream.size;
                }
            }
        }
    };

    CFile.prototype.findText = function(text, isMachingCase, isNext)
    {
        this.SearchResults.IsSearch = true;
        var pagesCount = this.pages.length;
        if (text === this.SearchResults.Text && isMachingCase === this.SearchResults.MachingCase)
        {
            if (this.SearchResults.Count === 0)
            {
                this.viewer.CurrentSearchNavi = null;
                this.SearchResults.CurrentPage = -1;
                this.SearchResults.Current = -1;
                return;
            }

            // поиск совпал, просто делаем навигацию к нужному месту
            if (isNext)
            {
                if ((this.SearchResults.Current + 1) < this.SearchResults.Pages[this.SearchResults.CurrentPage].length)
                {
                    // результат на этой же странице
                    this.SearchResults.Current++;
                }
                else
                {
                    var _pageFind = this.SearchResults.CurrentPage + 1;
                    var _bIsFound = false;
                    for (var i = _pageFind; i < pagesCount; i++)
                    {
                        if (0 < this.SearchResults.Pages[i].length)
                        {
                            this.SearchResults.Current = 0;
                            this.SearchResults.CurrentPage = i;
                            _bIsFound = true;
                            break;
                        }
                    }
                    if (!_bIsFound)
                    {
                        for (var i = 0; i < _pageFind; i++)
                        {
                            if (0 < this.SearchResults.Pages[i].length)
                            {
                                this.SearchResults.Current = 0;
                                this.SearchResults.CurrentPage = i;
                                _bIsFound = true;
                                break;
                            }
                        }
                    }
                }
            }
            else
            {
                if (this.SearchResults.Current > 0)
                {
                    // результат на этой же странице
                    this.SearchResults.Current--;
                }
                else
                {
                    var _pageFind = this.SearchResults.CurrentPage - 1;
                    var _bIsFound = false;
                    for (var i = _pageFind; i >= 0; i--)
                    {
                        if (0 < this.SearchResults.Pages[i].length)
                        {
                            this.SearchResults.Current = this.SearchResults.Pages[i].length - 1;
                            this.SearchResults.CurrentPage = i;
                            _bIsFound = true;
                            break;
                        }
                    }
                    if (!_bIsFound)
                    {
                        for (var i = pagesCount - 1; i > _pageFind; i--)
                        {
                            if (0 < this.SearchResults.Pages[i].length)
                            {
                                this.SearchResults.Current = this.SearchResults.Pages[i].length - 1;
                                this.SearchResults.CurrentPage = i;
                                _bIsFound = true;
                                break;
                            }
                        }
                    }
                }
            }

            this.viewer.CurrentSearchNavi = this.SearchResults.Pages[this.SearchResults.CurrentPage][this.SearchResults.Current];

            this.viewer.ToSearchResult();
            return;
        }
        // новый поиск
        for (var i = 0; i < this.pages.length; i++)
        {
            this.SearchResults.Pages[i].splice(0, this.SearchResults.Pages[i].length);
        }
        this.SearchResults.Count = 0;

        this.SearchResults.CurrentPage = -1;
        this.SearchResults.Current = -1;

        this.SearchResults.Text = text;
        this.SearchResults.MachingCase = isMachingCase;

        for (var i = 0; i < this.pages.length; i++)
        {
            this.searchPage(i);
            this.SearchResults.Count += this.SearchResults.Pages[i].length;
        }

        if (this.SearchResults.Count == 0)
        {
            this.viewer.CurrentSearchNavi = null;
            this.onUpdateOverlay();
            return;
        }

        var currentPage = this.viewer.currentPage;
        for (var i = currentPage; i < this.SearchResults.Pages.length; i++)
        {
            if (0 != this.SearchResults.Pages[i].length)
            {
                this.SearchResults.CurrentPage = i;
                this.SearchResults.Current = 0;
                break;
            }
        }
        if (this.SearchResults.Current === -1)
        {
            for (var i = 0; i < currentPage; i++)
            {
                if (0 != this.SearchResults.Pages[i].length)
                {
                    this.SearchResults.CurrentPage = i;
                    this.SearchResults.Current = 0;
                    break;
                }
            }
        }

        this.viewer.CurrentSearchNavi = this.SearchResults.Pages[this.SearchResults.CurrentPage][this.SearchResults.Current];
        this.viewer.ToSearchResult();
    };

    CFile.prototype.prepareSearch = function()
    {
        this.SearchResults.Pages = new Array(this.pages.length);
        for (var i = this.pages.length - 1; i >= 0; i--)
        {
            this.SearchResults.Pages[i] = [];
        }
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
            file.nativeFile["onRepaintPages"] = function(pages) {
                file.onRepaintPages && file.onRepaintPages(pages);
            };
            file.nativeFile["onUpdateStatistics"] = function(par, word, symbol, space) {
                file.onUpdateStatistics && file.onUpdateStatistics(par, word, symbol, space);
            };
            file.pages = file.nativeFile["getPages"]();

            for (var i = 0, len = file.pages.length; i < len; i++)
            {
                var page = file.pages[i];
                page.W = page["W"];
                page.H = page["H"];
                page.Dpi = page["Dpi"];
            }

            file.prepareSearch();
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
            file.nativeFile["onRepaintPages"] = function(pages) {
                file.onRepaintPages && file.onRepaintPages(pages);
            };
            file.nativeFile["onUpdateStatistics"] = function(par, word, symbol, space) {
                file.onUpdateStatistics && file.onUpdateStatistics(par, word, symbol, space);
            };
            file.pages = file.nativeFile["getPages"]();

            for (var i = 0, len = file.pages.length; i < len; i++)
            {
                var page = file.pages[i];
                page.W = page["W"];
                page.H = page["H"];
                page.Dpi = page["Dpi"];
            }

            file.prepareSearch();
            //file.cacheManager = new AscCommon.CCacheManager();
        }
    };

})(window, undefined);
