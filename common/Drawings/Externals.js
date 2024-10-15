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

(function(window, document){

    // Import
    var FontStyle = AscFonts.FontStyle;

    // глобальные мапы для быстрого поиска
    var g_map_font_index = {};
    var g_fonts_streams = [];

    var isLoadFontsSync = (window["NATIVE_EDITOR_ENJINE"] === true);

    function CFontFileLoader(id)
    {
        this.LoadingCounter = 0;
        this.Id             = id;
        this.Status         = -1;  // -1 - notloaded, 0 - loaded, 1 - error, 2 - loading
        this.stream_index   = -1;
        this.callback       = null;
    }

    CFontFileLoader.prototype.GetMaxLoadingCount = function()
    {
        return 3;
    };
    CFontFileLoader.prototype.CheckLoaded = function()
    {
        return (0 === this.Status || 1 === this.Status);
    };
    CFontFileLoader.prototype.SetStreamIndex = function(index)
    {
        this.stream_index = index;
    };
    CFontFileLoader.prototype.LoadFontFromData = function(data)
    {
        let stream_index = g_fonts_streams.length;
        g_fonts_streams[stream_index] = new AscFonts.FontStream(data, data.length);
        this.SetStreamIndex(stream_index);
        this.Status = 0;
    };
    CFontFileLoader.prototype.LoadFontNative = function()
    {
        let data = window["native"]["GetFontBinary"](this.Id);
        this.LoadFontFromData(data);
    };
    CFontFileLoader.prototype.LoadFontArrayBuffer = function(basePath)
    {
        var xhr = new XMLHttpRequest();
        xhr.fontFile = this;
        xhr.open('GET', basePath + this.Id, true);

        if (typeof ArrayBuffer !== 'undefined' && !window.opera)
            xhr.responseType = 'arraybuffer';

        if (xhr.overrideMimeType)
            xhr.overrideMimeType('text/plain; charset=x-user-defined');
        else
            xhr.setRequestHeader('Accept-Charset', 'x-user-defined');

        xhr.onload = function()
        {
            if (this.status !== 200)
                return this.onerror();

            this.fontFile.Status = 0;
            if (typeof ArrayBuffer !== 'undefined' && !window.opera && this.response)
            {
                this.fontFile.LoadFontFromData(new Uint8Array(this.response));
            }
            else if (AscCommon.AscBrowser.isIE)
            {
                let _response = new VBArray(this["responseBody"]).toArray();

                let srcLen = _response.length;
                let stream = new AscFonts.FontStream(AscFonts.allocate(srcLen), srcLen);

                let dstPx = stream.data;
                let index = 0;

                while (index < srcLen)
                {
                    dstPx[index] = _response[index];
                    index++;
                }

                let stream_index = g_fonts_streams.length;
                g_fonts_streams[stream_index] = stream;
                this.fontFile.SetStreamIndex(stream_index);
            }
            else
            {
                let stream_index = g_fonts_streams.length;
                g_fonts_streams[stream_index] = AscFonts.CreateFontData3(this.responseText);
                this.fontFile.SetStreamIndex(stream_index);
            }

            // decode
            let guidOdttf = [0xA0, 0x66, 0xD6, 0x20, 0x14, 0x96, 0x47, 0xfa, 0x95, 0x69, 0xB8, 0x50, 0xB0, 0x41, 0x49, 0x48];
            let stream = g_fonts_streams[g_fonts_streams.length - 1];
            let data = stream.data;

            let count_decode = Math.min(32, stream.size);
            for (let i = 0; i < count_decode; ++i)
                data[i] ^= guidOdttf[i % 16];

            if (null != this.fontFile.callback)
                this.fontFile.callback();
            if (this.fontFile["externalCallback"])
                this.fontFile["externalCallback"]();
        };
        xhr.onerror = function()
        {
            this.fontFile.LoadingCounter++;
            if (this.fontFile.LoadingCounter < this.fontFile.GetMaxLoadingCount())
            {
                //console.log("font loaded: one more attemption");
                this.fontFile.Status = -1;
                return;
            }

            this.fontFile.Status = 2; // aka loading...
            window["Asc"]["editor"].sendEvent("asc_onError", Asc.c_oAscError.ID.LoadingFontError, Asc.c_oAscError.Level.Critical);
        };

        xhr.send(null);
    };
    CFontFileLoader.prototype.LoadFontAsync = function(basePath, callback)
    {
        if (isLoadFontsSync)
            return true;

        if (window["AscDesktopEditor"] !== undefined)
        {
            if (-1 !== this.Status)
                return true;

            window["AscDesktopEditor"]["LoadFontBase64"](this.Id);

            let streams_count = g_fonts_streams.length;
            g_fonts_streams[streams_count] = AscFonts.CreateFontData4(window[this.Id]);
            this.SetStreamIndex(streams_count);

            this.Status = 0;
            delete window[this.Id];

            if (callback)
                callback();
            if (this["externalCallback"])
                this["externalCallback"]();
            return;
        }

        this.callback = callback;

        if (-1 !== this.Status)
            return true;

        this.Status = 2;
        this.LoadFontArrayBuffer(basePath);
        return false;
    };

    CFontFileLoader.prototype["LoadFontAsync"] = CFontFileLoader.prototype.LoadFontAsync;
    CFontFileLoader.prototype["GetID"] = function() { return this.Id; };
    CFontFileLoader.prototype["GetStatus"] = function() { return this.Status; };
    CFontFileLoader.prototype["GetStreamIndex"] = function() { return this.stream_index; };

    function CFontFileLoaderEmbed(id)
    {
        this.Id             = id;
        this.Status         = -1;  // -1 - notloaded, 0 - loaded, 1 - error, 2 - loading
        this.stream_index   = -1;
        this.callback       = null;
    }

    CFontFileLoaderEmbed.prototype.startIndex = -1;
    CFontFileLoaderEmbed.prototype.CheckLoaded = function()
    {
        return (0 === this.Status || 1 === this.Status);
    };
    CFontFileLoaderEmbed.prototype.SetStreamIndex = function(index)
    {
        this.stream_index = index;
    };
    CFontFileLoaderEmbed.prototype.LoadFontAsync = function()
    {
        if (this.Status === 0 || this.Status === 1)
            return true;

        let fontData = AscFonts.loadEmbeddedFont(this.Id);
        if (fontData)
        {
            let stream_index = g_fonts_streams.length;
            g_fonts_streams[stream_index] = new AscFonts.FontStream(fontData, fontData.length);
            this.SetStreamIndex(stream_index);
            this.Status = 0;
            AscFonts.CreateNativeStreamByIndex(stream_index);
        }
        else
        {
            this.Status = 1;
        }

        return true;
    };

    const fontstyle_mask_regular    = 1;
    const fontstyle_mask_italic     = 2;
    const fontstyle_mask_bold       = 4;
    const fontstyle_mask_bolditalic = 8;

    function GenerateMapId(api, name, style, size)
    {
        var fontInfo = api.FontLoader.fontInfos[api.FontLoader.map_font_index[name]];
        let info = fontInfo.GetNeedInfo(style);

        var _ext = "";
        if (info.needB)
            _ext += "nbold";
        if (info.needI)
            _ext += "nitalic";

        // index != -1 (!!!)
        let fontfile = api.FontLoader.fontFiles[info.index];
        return fontfile.Id + info.faceIndex + size + _ext;
    }

    function CFontInfo(sName, thumbnail, indexR, faceIndexR, indexI, faceIndexI, indexB, faceIndexB, indexBI, faceIndexBI)
    {
        this.Name = sName;
        this.Thumbnail = thumbnail;
        this.NeedStyles = 0;

        this.indexR     = indexR;
        this.faceIndexR = faceIndexR;
        this.needR      = false;

        this.indexI     = indexI;
        this.faceIndexI = faceIndexI;
        this.needI      = false;

        this.indexB     = indexB;
        this.faceIndexB = faceIndexB;
        this.needB      = false;

        this.indexBI    = indexBI;
        this.faceIndexBI= faceIndexBI;
        this.needBI     = false;
    }

    CFontInfo.prototype =
    {
        // начинаем грузить нужные стили
        CheckFontLoadStyles : function(global_loader)
        {
            if (isLoadFontsSync)
                return false;

            if ((this.NeedStyles & 0x0F) === 0x0F)
            {
                this.needR = true;
                this.needI = true;
                this.needB = true;
                this.needBI = true;
            }
            else
            {
                let needs = [false, false, false, false];
                if ((this.NeedStyles & fontstyle_mask_regular) !== 0)
                    needs[this.GetBaseStyle(FontStyle.FontStyleRegular)] = true;
                if ((this.NeedStyles & fontstyle_mask_italic) !== 0)
                    needs[this.GetBaseStyle(FontStyle.FontStyleItalic)] = true;
                if ((this.NeedStyles & fontstyle_mask_bold) !== 0)
                    needs[this.GetBaseStyle(FontStyle.FontStyleBold)] = true;
                if ((this.NeedStyles & fontstyle_mask_bolditalic) !== 0)
                    needs[this.GetBaseStyle(FontStyle.FontStyleBoldItalic)] = true;

                if (needs[0]) this.needR = true;
                if (needs[1]) this.needI = true;
                if (needs[2]) this.needB = true;
                if (needs[3]) this.needBI = true;
            }

            var fonts = global_loader.fontFiles;
            var basePath = global_loader.fontFilesPath;
            var isNeed = false;
            if ((this.needR === true) && (-1 !== this.indexR) && (fonts[this.indexR].CheckLoaded() === false))
            {
                fonts[this.indexR].LoadFontAsync(basePath, null);
                isNeed = true;
            }
            if ((this.needI === true) && (-1 !== this.indexI) && (fonts[this.indexI].CheckLoaded() === false))
            {
                fonts[this.indexI].LoadFontAsync(basePath, null);
                isNeed = true;
            }
            if ((this.needB === true) && (-1 !== this.indexB) && (fonts[this.indexB].CheckLoaded() === false))
            {
                fonts[this.indexB].LoadFontAsync(basePath, null);
                isNeed = true;
            }
            if ((this.needBI === true) && (-1 !== this.indexBI) && (fonts[this.indexBI].CheckLoaded() === false))
            {
                fonts[this.indexBI].LoadFontAsync(basePath, null);
                isNeed = true;
            }

            return isNeed;
        },

        // надо ли грузить хоть один шрифт из семейства
        CheckFontLoadStylesNoLoad : function(global_loader)
        {
            if (isLoadFontsSync)
                return false;
            var fonts = global_loader.fontFiles;
            if ((-1 !== this.indexR) && (fonts[this.indexR].CheckLoaded() === false))
                return true;
            if ((-1 !== this.indexI) && (fonts[this.indexI].CheckLoaded() === false))
                return true;
            if ((-1 !== this.indexB) && (fonts[this.indexB].CheckLoaded() === false))
                return true;
            if ((-1 !== this.indexBI) && (fonts[this.indexBI].CheckLoaded() === false))
                return true;
            return false;
        },

        // используется только в тестовом примере
        LoadFontsFromServer : function(global_loader)
        {
            var fonts = global_loader.fontFiles;
            var basePath = global_loader.fontFilesPath;
            if ((-1 !== this.indexR) && (fonts[this.indexR].CheckLoaded() === false))
                fonts[this.indexR].LoadFontAsync(basePath, null);
            if ((-1 !== this.indexI) && (fonts[this.indexI].CheckLoaded() === false))
                fonts[this.indexI].LoadFontAsync(basePath, null);
            if ((-1 !== this.indexB) && (fonts[this.indexB].CheckLoaded() === false))
                fonts[this.indexB].LoadFontAsync(basePath, null);
            if ((-1 !== this.indexBI) && (fonts[this.indexBI].CheckLoaded() === false))
                fonts[this.indexBI].LoadFontAsync(basePath, null);
        },

        LoadFont : function(font_loader, fontManager, fEmSize, style, dHorDpi, dVerDpi, transform, isNoSetupToManager)
        {
            let info = this.GetNeedInfo(style);
            let fontfile = font_loader.fontFiles[info.index];

            if (window["NATIVE_EDITOR_ENJINE"] && fontfile.Status !== 0)
                fontfile.LoadFontNative();

            var pFontFile = fontManager.LoadFont(fontfile, info.faceIndex, fEmSize,
                (0 !== (style & FontStyle.FontStyleBold)) ? true : false,
                (0 !== (style & FontStyle.FontStyleItalic)) ? true : false,
                info.needB, info.needI,
                isNoSetupToManager);

            if (!pFontFile && -1 === fontfile.stream_index && true === AscFonts.IsLoadFontOnCheckSymbols && true != AscFonts.IsLoadFontOnCheckSymbolsWait)
            {
                // в форматах pdf/xps - не прогоняем символы через чеккер при открытии,
                // так как там должны быть символы в встроенном шрифте. Но вдруг?
                // тогда при отрисовке СРАЗУ грузим шрифт - и при загрузке перерисовываемся
                // сюда попали только если символ попал в чеккер
                AscFonts.IsLoadFontOnCheckSymbols = false;
                AscFonts.IsLoadFontOnCheckSymbolsWait = true;
                AscFonts.FontPickerByCharacter.loadFonts(window.editor, function ()
                {
                    AscFonts.IsLoadFontOnCheckSymbolsWait = false;
                    this.WordControl && this.WordControl.private_RefreshAll();
                });
            }

            if (pFontFile && (true !== isNoSetupToManager))
            {
                var newEmSize = fontManager.UpdateSize(fEmSize, dVerDpi, dVerDpi);
                pFontFile.SetSizeAndDpi(newEmSize, dHorDpi, dVerDpi);

                if (undefined !== transform)
                {
                    fontManager.SetTextMatrix2(transform.sx,transform.shy,transform.shx,transform.sy,transform.tx,transform.ty);
                }
                else
                {
                    fontManager.SetTextMatrix(1, 0, 0, 1, 0, 0);
                }
            }

            return pFontFile;
        },

        GetFontID : function(font_loader, style)
        {
            let info = this.GetNeedInfo(style);
            let fontfile = font_loader.fontFiles[info.index];
            return { id: fontfile.Id, faceIndex : info.faceIndex, file : fontfile };
        },

        // по запрашиваемому стилю - отдаем какой будем использовать
        GetBaseStyle : function(style)
        {
            switch (style)
            {
                case FontStyle.FontStyleBoldItalic:
                {
                    if (-1 !== this.indexBI)
                        return FontStyle.FontStyleBoldItalic;
                    else if (-1 !== this.indexB)
                        return FontStyle.FontStyleBold;
                    else if (-1 !== this.indexI)
                        return FontStyle.FontStyleItalic;
                    else
                        return FontStyle.FontStyleRegular;
                    break;
                }
                case FontStyle.FontStyleBold:
                {
                    if (-1 !== this.indexB)
                        return FontStyle.FontStyleBold;
                    else if (-1 !== this.indexR)
                        return FontStyle.FontStyleRegular;
                    else if (-1 !== this.indexBI)
                        return FontStyle.FontStyleBoldItalic;
                    else
                        return FontStyle.FontStyleItalic;
                    break;
                }
                case FontStyle.FontStyleItalic:
                {
                    if (-1 !== this.indexI)
                        return FontStyle.FontStyleItalic;
                    else if (-1 !== this.indexR)
                        return FontStyle.FontStyleRegular;
                    else if (-1 !== this.indexBI)
                        return FontStyle.FontStyleBoldItalic;
                    else
                        return FontStyle.FontStyleBold;
                    break;
                }
                case FontStyle.FontStyleRegular:
                {
                    if (-1 !== this.indexR)
                        return FontStyle.FontStyleRegular;
                    else if (-1 !== this.indexI)
                        return FontStyle.FontStyleItalic;
                    else if (-1 !== this.indexB)
                        return FontStyle.FontStyleBold;
                    else
                        return FontStyle.FontStyleBoldItalic;
                }
            }
            return FontStyle.FontStyleRegular;
        },

        // по запрашиваемому стилю - возвращаем какой будем грузить и какие настройки нужно сделать самому
        GetNeedInfo : function(style)
        {
            let result = {
                index : -1,
                faceIndex : 0,
                needB : false,
                needI : false
            };

            let resStyle = this.GetBaseStyle(style);

            switch (resStyle)
            {
                case FontStyle.FontStyleBoldItalic:
                {
                    result.index = this.indexBI;
                    result.faceIndex = this.faceIndexBI;
                    break;
                }
                case FontStyle.FontStyleBold:
                {
                    result.index = this.indexB;
                    result.faceIndex = this.faceIndexB;
                    if (0 !== (style & FontStyle.FontStyleItalic))
                        result.needI = true;
                    break;
                }
                case FontStyle.FontStyleItalic:
                {
                    result.index = this.indexI;
                    result.faceIndex = this.faceIndexI;
                    if (0 !== (style & FontStyle.FontStyleBold))
                        result.needB = true;
                    break;
                }
                case FontStyle.FontStyleRegular:
                default:
                {
                    result.index = this.indexR;
                    result.faceIndex = this.faceIndexR;
                    if (0 !== (style & FontStyle.FontStyleItalic))
                        result.needI = true;
                    if (0 !== (style & FontStyle.FontStyleBold))
                        result.needB = true;
                    break;
                }
            }

            return result;
        }
    };

    function CFontInfoEmbed(name, index)
    {
        this.Name = name;
        this.indexR     = index;
        this.faceIndexR = 0;
    }

    CFontInfoEmbed.prototype =
    {
        // начинаем грузить нужные стили
        CheckFontLoadStyles : function(global_loader)
        {
            let fontFile = global_loader.fontFiles[this.indexR];
            fontFile.LoadFontAsync();
            return false;
        },

        // надо ли грузить хоть один шрифт из семейства
        CheckFontLoadStylesNoLoad : function(global_loader)
        {
            return false;
        },

        LoadFont : function(font_loader, fontManager, fEmSize, style, dHorDpi, dVerDpi, transform, isNoSetupToManager)
        {
            let fontfile = font_loader.fontFiles[this.indexR];

            var pFontFile = fontManager.LoadFont(fontfile, this.faceIndexR, fEmSize,
                false,
                false,
                false, false,
                isNoSetupToManager);

            if (pFontFile && (true !== isNoSetupToManager))
            {
                pFontFile.m_pFaceInfo.family_name = this.Name;

                var newEmSize = fontManager.UpdateSize(fEmSize, dVerDpi, dVerDpi);
                pFontFile.SetSizeAndDpi(newEmSize, dHorDpi, dVerDpi);

                if (undefined !== transform)
                {
                    fontManager.SetTextMatrix2(transform.sx,transform.shy,transform.shx,transform.sy,transform.tx,transform.ty);
                }
                else
                {
                    fontManager.SetTextMatrix(1, 0, 0, 1, 0, 0);
                }
            }

            return pFontFile;
        },

        GetFontID : function(font_loader)
        {
            let fontfile = font_loader.fontFiles[this.indexК];
            return { id: fontfile.Id, faceIndex : this.faceIndexR, file : fontfile };
        }
    };

    // thumbnail - это позиция (y) в общем табнейле всех шрифтов
    function CFont(name, id, thumbnail, style)
    {
        this.name       = name;
        this.id         = id || "";
        this.thumbnail  = thumbnail || 0;
        this.NeedStyles = style || (fontstyle_mask_regular | fontstyle_mask_italic | fontstyle_mask_bold | fontstyle_mask_bolditalic);
    }

    CFont.prototype["asc_getFontId"] = CFont.prototype.asc_getFontId = function() { return this.id; };
    CFont.prototype["asc_getFontName"] = CFont.prototype.asc_getFontName = function()
    {
        var _name = AscFonts.g_fontApplication ? AscFonts.g_fontApplication.NameToInterface[this.name] : null;
        return _name ? _name : this.name;
    };
    CFont.prototype["asc_getFontThumbnail"] = CFont.prototype.asc_getFontThumbnail = function() { return this.thumbnail; };
    // для совместимости
    CFont.prototype["asc_getFontType"] = CFont.prototype.asc_getFontType = function() { return 1; };

    var ImageLoadStatus =
    {
        Loading : 0,
        Complete : 1
    };

    function CImage(src)
    {
        this.src    = src;
        this.Image  = null;
        this.Status = ImageLoadStatus.Complete;
    }

	function checkAllFonts()
    {
        if (undefined === window["__fonts_files"] && window["native"] && window["native"]["GenerateAllFonts"])
            window["native"]["GenerateAllFonts"]();

        if (undefined === window["__fonts_files"])
            return;

        let g_font_files, g_font_infos;

        var files = window["__fonts_files"];

		let count_files = files ? files.length : 0;
		g_font_files = new Array(count_files);
		for (let i = 0; i < count_files; i++)
		{
			g_font_files[i] = new CFontFileLoader(files[i]);
		}

		let infos = window["__fonts_infos"];
		let count_infos = infos ? infos.length : 0;
		g_font_infos = new Array(count_infos);

		let curIndex = 0;
		let ascFontFath = "ASC.ttf";
		for (let i = 0; i < count_infos; i++)
		{
			let info = infos[i];
			if ("ASCW3" === info[0])
			{
			    ascFontFath = g_font_files[info[1]].Id;
                continue;
            }

			g_font_infos[curIndex] = new CFontInfo(info[0], i, info[1], info[2], info[3], info[4], info[5], info[6], info[7], info[8]);
			g_map_font_index[info[0]] = curIndex;
			curIndex++;
		}

        g_font_infos.length = curIndex;

		/////////////////////////////////////////////////////////////////////
		// наш шрифт для спецсимволов
		let ascW3 = new CFontFileLoader(ascFontFath);
        ascW3.Status = 0;
		let streams_count = g_fonts_streams.length;
		g_fonts_streams[streams_count] = AscFonts.CreateFontData2("AAEAAAARAQAABAAQTFRTSOGXTFYAAAIgAAAADk9TLzJ0/uscAAABmAAAAGBWRE1Yblp14wAAAjAAAAXgY21hcNCBcnoAAAksAAAAaGN2dCBBCTniAAAUWAAAAohmcGdtu26+2AAACZQAAAi4Z2FzcAAPABYAAB9cAAAADGdseWaWCCxJAAAW4AAABQBoZG1466dcHQAACBAAAAEcaGVhZBMko9YAAAEcAAAANmhoZWEOeQN0AAABVAAAACRobXR4MfIEPAAAAfgAAAAobG9jYQVeBDYAABvgAAAAFm1heHAIXQkWAAABeAAAACBuYW1lZvQXkQAAG/gAAALocG9zdKnSYBIAAB7gAAAAenByZXDMvTRHAAASTAAAAgsAAQAAAAEAAFPp5kFfDzz1ABsIAAAAAADPTriwAAAAAN2yoHYAgP/nBnUFyAAAAAwAAQAAAAAAAAABAAAHPv5OAEMHIQCA/foGdQABAAAAAAAAAAAAAAAAAAAACgABAAAACgBIAAMAAAAAAAIAEAAUAFcAAAfoCLgAAAAAAAMFjAGQAAUAAATOBM4AAAMWBM4EzgAAAxYAZgISDAAFBAECAQgHBwcHAAAAAAAAAAAAAAAAAAAAAE1TICAAQCXJ8DgF0/5RATMHPgGygAAAAAAAAAAEJgW7AAAAIAAABAAAgAAAAAAE0gAAAgAAAAchAK0EbwCtBuQAjAbkAIwG5AClBuQApQAAAApLAQEBS0tLS0tLAAAAAAABAAEBAQEBAAwA+Aj/AAgACP/+AAkACf/+AAoACv/9AAsACv/9AAwAC//9AA0ADP/9AA4ADf/9AA8ADv/8ABAAD//8ABEAEP/8ABIAEf/8ABMAEv/7ABQAE//7ABUAFP/7ABYAFP/7ABcAFf/7ABgAFv/6ABkAF//6ABoAGP/6ABsAGf/6ABwAGv/6AB0AG//5AB4AHP/5AB8AHf/5ACAAHf/5ACEAHv/5ACIAH//4ACMAIP/4ACQAIf/4ACUAIv/4ACYAI//3ACcAJP/3ACgAJf/3ACkAJv/3ACoAJ//3ACsAJ//2ACwAKP/2AC0AKf/2AC4AKv/2AC8AK//2ADAALP/1ADEALf/1ADIALv/1ADMAL//1ADQAMP/0ADUAMP/0ADYAMf/0ADcAMv/0ADgAM//0ADkANP/zADoANf/zADsANv/zADwAN//zAD0AOP/zAD4AOf/yAD8AOv/yAEAAOv/yAEEAO//yAEIAPP/yAEMAPf/xAEQAPv/xAEUAP//xAEYAQP/xAEcAQf/wAEgAQv/wAEkAQ//wAEoAQ//wAEsARP/wAEwARf/vAE0ARv/vAE4AR//vAE8ASP/vAFAASf/vAFEASv/uAFIAS//uAFMATP/uAFQATf/uAFUATf/tAFYATv/tAFcAT//tAFgAUP/tAFkAUf/tAFoAUv/sAFsAU//sAFwAVP/sAF0AVf/sAF4AVv/sAF8AV//rAGAAV//rAGEAWP/rAGIAWf/rAGMAWv/rAGQAW//qAGUAXP/qAGYAXf/qAGcAXv/qAGgAX//pAGkAYP/pAGoAYP/pAGsAYf/pAGwAYv/pAG0AY//oAG4AZP/oAG8AZf/oAHAAZv/oAHEAZ//oAHIAaP/nAHMAaf/nAHQAav/nAHUAav/nAHYAa//mAHcAbP/mAHgAbf/mAHkAbv/mAHoAb//mAHsAcP/lAHwAcf/lAH0Acv/lAH4Ac//lAH8Ac//lAIAAdP/kAIEAdf/kAIIAdv/kAIMAd//kAIQAeP/kAIUAef/jAIYAev/jAIcAe//jAIgAfP/jAIkAff/iAIoAff/iAIsAfv/iAIwAf//iAI0AgP/iAI4Agf/hAI8Agv/hAJAAg//hAJEAhP/hAJIAhf/hAJMAhv/gAJQAhv/gAJUAh//gAJYAiP/gAJcAif/gAJgAiv/fAJkAi//fAJoAjP/fAJsAjf/fAJwAjv/eAJ0Aj//eAJ4AkP/eAJ8AkP/eAKAAkf/eAKEAkv/dAKIAk//dAKMAlP/dAKQAlf/dAKUAlv/dAKYAl//cAKcAmP/cAKgAmf/cAKkAmf/cAKoAmv/bAKsAm//bAKwAnP/bAK0Anf/bAK4Anv/bAK8An//aALAAoP/aALEAof/aALIAov/aALMAo//aALQAo//ZALUApP/ZALYApf/ZALcApv/ZALgAp//ZALkAqP/YALoAqf/YALsAqv/YALwAq//YAL0ArP/XAL4Arf/XAL8Arf/XAMAArv/XAMEAr//XAMIAsP/WAMMAsf/WAMQAsv/WAMUAs//WAMYAtP/WAMcAtf/VAMgAtv/VAMkAtv/VAMoAt//VAMsAuP/UAMwAuf/UAM0Auv/UAM4Au//UAM8AvP/UANAAvf/TANEAvv/TANIAv//TANMAwP/TANQAwP/TANUAwf/SANYAwv/SANcAw//SANgAxP/SANkAxf/SANoAxv/RANsAx//RANwAyP/RAN0Ayf/RAN4Ayf/QAN8Ayv/QAOAAy//QAOEAzP/QAOIAzf/QAOMAzv/PAOQAz//PAOUA0P/PAOYA0f/PAOcA0v/PAOgA0//OAOkA0//OAOoA1P/OAOsA1f/OAOwA1v/NAO0A1//NAO4A2P/NAO8A2f/NAPAA2v/NAPEA2//MAPIA3P/MAPMA3P/MAPQA3f/MAPUA3v/MAPYA3//LAPcA4P/LAPgA4f/LAPkA4v/LAPoA4//LAPsA5P/KAPwA5f/KAP0A5v/KAP4A5v/KAP8A5//JAAAAFwAAAAwJCAUABQIIBQgICAgKCQUABgMJBgkJCQkLCgYABwMKBgkJCQkMCwYABwMLBwoKCgoNDAcACAMMBwsLCwsPDQgACQQNCA0NDQ0QDggACgQOCQ4ODg4RDwkACgQPCQ8PDw8TEQoACwURCxAQEBAVEwsADQUTDBISEhIYFQwADgYVDRUVFRUbGA4AEAcYDxcXFxcdGg8AEQcaEBkZGRkgHRAAEwgdEhwcHBwhHREAFAgdEhwcHBwlIRMAFgkhFSAgICAqJRUAGQslFyQkJCQuKRcAHAwpGigoKCgyLRkAHg0tHCsrKys2MBsAIQ4wHi8vLy86NB0AIw80IDIyMjJDPCIAKBE8JTo6OjpLQyYALRNDKkFBQUEAAAACAAEAAAAAABQAAwAAAAAAIAAGAAwAAP//AAEAAAAEAEgAAAAOAAgAAgAGJcklyyYR8CDwIvA4//8AACXJJcsmEPAg8CLwOP//2j3aPNn4D+MP4g/NAAEAAAAAAAAAAAAAAAAAAEA3ODc0MzIxMC8uLSwrKikoJyYlJCMiISAfHh0cGxoZGBcWFRQTEhEQDw4NDAsKCQgHBgUEAwIBACxFI0ZgILAmYLAEJiNISC0sRSNGI2EgsCZhsAQmI0hILSxFI0ZgsCBhILBGYLAEJiNISC0sRSNGI2GwIGAgsCZhsCBhsAQmI0hILSxFI0ZgsEBhILBmYLAEJiNISC0sRSNGI2GwQGAgsCZhsEBhsAQmI0hILSwBECA8ADwtLCBFIyCwzUQjILgBWlFYIyCwjUQjWSCw7VFYIyCwTUQjWSCwBCZRWCMgsA1EI1khIS0sICBFGGhEILABYCBFsEZ2aIpFYEQtLAGxCwpDI0NlCi0sALEKC0MjQwstLACwRiNwsQFGPgGwRiNwsQJGRTqxAgAIDS0sRbBKI0RFsEkjRC0sIEWwAyVFYWSwUFFYRUQbISFZLSywAUNjI2KwACNCsA8rLSwgRbAAQ2BELSwBsAZDsAdDZQotLCBpsEBhsACLILEswIqMuBAAYmArDGQjZGFcWLADYVktLEWwESuwRyNEsEd65BgtLLgBplRYsAlDuAEAVFi5AEr/gLFJgEREWVktLLASQ1iHRbARK7AXI0SwF3rkGwOKRRhpILAXI0SKiocgsKBRWLARK7AXI0SwF3rkGyGwF3rkWVkYLSwtLEtSWCFFRBsjRYwgsAMlRVJYRBshIVlZLSwBGC8tLCCwAyVFsEkjREWwSiNERWUjRSCwAyVgaiCwCSNCI2iKamBhILAairAAUnkhshpKQLn/4ABKRSCKVFgjIbA/GyNZYUQcsRQAilJ5s0lAIElFIIpUWCMhsD8bI1lhRC0ssRARQyNDCy0ssQ4PQyNDCy0ssQwNQyNDCy0ssQwNQyNDZQstLLEOD0MjQ2ULLSyxEBFDI0NlCy0sS1JYRUQbISFZLSwBILADJSNJsEBgsCBjILAAUlgjsAIlOCOwAiVlOACKYzgbISEhISFZAS0sRWmwCUNgihA6LSwBsAUlECMgivUAsAFgI+3sLSwBsAUlECMgivUAsAFhI+3sLSwBsAYlEPUA7ewtLCCwAWABECA8ADwtLCCwAWEBECA8ADwtLLArK7AqKi0sALAHQ7AGQwstLD6wKiotLDUtLHawSyNwECCwS0UgsABQWLABYVk6LxgtLCEhDGQjZIu4QABiLSwhsIBRWAxkI2SLuCAAYhuyAEAvK1mwAmAtLCGwwFFYDGQjZIu4FVViG7IAgC8rWbACYC0sDGQjZIu4QABiYCMhLSy0AAEAAAAVsAgmsAgmsAgmsAgmDxAWE0VoOrABFi0stAABAAAAFbAIJrAIJrAIJrAIJg8QFhNFaGU6sAEWLSxFIyBFILEEBSWKUFgmYYqLGyZgioxZRC0sRiNGYIqKRiMgRopgimG4/4BiIyAQI4qxS0uKcEVgILAAUFiwAWG4/8CLG7BAjFloATotLLAzK7AqKi0ssBNDWAMbAlktLLATQ1gCGwNZLbgAOSxLuAAMUFixAQGOWbgB/4W4AEQduQAMAANfXi24ADosICBFaUSwAWAtuAA7LLgAOiohLbgAPCwgRrADJUZSWCNZIIogiklkiiBGIGhhZLAEJUYgaGFkUlgjZYpZLyCwAFNYaSCwAFRYIbBAWRtpILAAVFghsEBlWVk6LbgAPSwgRrAEJUZSWCOKWSBGIGphZLAEJUYgamFkUlgjilkv/S24AD4sSyCwAyZQWFFYsIBEG7BARFkbISEgRbDAUFiwwEQbIVlZLbgAPywgIEVpRLABYCAgRX1pGESwAWAtuABALLgAPyotuABBLEsgsAMmU1iwQBuwAFmKiiCwAyZTWCMhsICKihuKI1kgsAMmU1gjIbgAwIqKG4ojWSCwAyZTWCMhuAEAioobiiNZILADJlNYIyG4AUCKihuKI1kguAADJlNYsAMlRbgBgFBYIyG4AYAjIRuwAyVFIyEjIVkbIVlELbgAQixLU1hFRBshIVktuABDLEu4AAxQWLEBAY5ZuAH/hbgARB25AAwAA19eLbgARCwgIEVpRLABYC24AEUsuABEKiEtuABGLCBGsAMlRlJYI1kgiiCKSWSKIEYgaGFksAQlRiBoYWRSWCNlilkvILAAU1hpILAAVFghsEBZG2kgsABUWCGwQGVZWTotuABHLCBGsAQlRlJYI4pZIEYgamFksAQlRiBqYWRSWCOKWS/9LbgASCxLILADJlBYUViwgEQbsEBEWRshISBFsMBQWLDARBshWVktuABJLCAgRWlEsAFgICBFfWkYRLABYC24AEosuABJKi24AEssSyCwAyZTWLBAG7AAWYqKILADJlNYIyGwgIqKG4ojWSCwAyZTWCMhuADAioobiiNZILADJlNYIyG4AQCKihuKI1kgsAMmU1gjIbgBQIqKG4ojWSC4AAMmU1iwAyVFuAGAUFgjIbgBgCMhG7ADJUUjISMhWRshWUQtuABMLEtTWEVEGyEhWS24AE0sS7gACVBYsQEBjlm4Af+FuABEHbkACQADX14tuABOLCAgRWlEsAFgLbgATyy4AE4qIS24AFAsIEawAyVGUlgjWSCKIIpJZIogRiBoYWSwBCVGIGhhZFJYI2WKWS8gsABTWGkgsABUWCGwQFkbaSCwAFRYIbBAZVlZOi24AFEsIEawBCVGUlgjilkgRiBqYWSwBCVGIGphZFJYI4pZL/0tuABSLEsgsAMmUFhRWLCARBuwQERZGyEhIEWwwFBYsMBEGyFZWS24AFMsICBFaUSwAWAgIEV9aRhEsAFgLbgAVCy4AFMqLbgAVSxLILADJlNYsEAbsABZioogsAMmU1gjIbCAioobiiNZILADJlNYIyG4AMCKihuKI1kgsAMmU1gjIbgBAIqKG4ojWSCwAyZTWCMhuAFAioobiiNZILgAAyZTWLADJUW4AYBQWCMhuAGAIyEbsAMlRSMhIyFZGyFZRC24AFYsS1NYRUQbISFZLbgATSsBugACAUAATysBvwFBAFsASwA6ACoAGQAAAFUrAL8BQABbAEsAOgAqABkAAABVKwC6AUIAAQBUK7gBPyBFfWkYRLgAQysAugE9AAEASiu4ATwgRX1pGES4ADkrAboBOQABADsrAb8BOQBNAD8AMQAjABUAAABBKwC6AToAAQBAK7gBOCBFfWkYREAMAEZGAAAAEhEIQEgguAEcskgyILgBA0BhSDIgv0gyIIlIMiCHSDIghkgyIGdIMiBlSDIgYUgyIFxIMiBVSDIgiEgyIGZIMiBiSDIgYEgyN5BqByQIIgggCB4IHAgaCBgIFggUCBIIEAgOCAwICggICAYIBAgCCAAIALATA0sCS1NCAUuwwGMAS2IgsPZTI7gBClFasAUjQgGwEksAS1RCGLkAAQfAhY2wOCuwAoi4AQBUWLgB/7EBAY6FG7ASQ1i5AAEB/4WNG7kAAQH/hY1ZWQAWKysrKysrKysrKysrKysrKysrKxgrGCsBslAAMkthi2AdACsrKysBKysrKysrKysrKysBRWlTQgFLUFixCABCWUNcWLEIAEJZswILChJDWGAbIVlCFhBwPrASQ1i5OyEYfhu6BAABqAALK1mwDCNCsA0jQrASQ1i5LUEtQRu6BAAEAAALK1mwDiNCsA8jQrASQ1i5GH47IRu6AagEAAALK1mwECNCsBEjQgEAAAAAAAAAAAAAAAAAEDgF4gAAAAAA7gAEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAP///////////////////////////////wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAD///////8AAAAAAGMAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAYwCUAJQAlACUBcgFyABiAJQAlAXIAGIAYgBjAJQBUQBjAGMAgACUAYoCTwLkBcgAlACUAJQAlAC6APcBKAEoASgBWQHuAh8FyADFAMYBKAEoBEsEVgRWBS8FyAXIBisAMQBiAGIAYwBjAJQAlACUAJQAlADFAMYAxgDeAPYA9wD3APcBKAEoASgBKAFZAXIBcgGKAYoBvAHtAe0B7gJQAlACUAJRApoCmgLkAuQDTQOzBEsESwRWBKAEqwSrBQIFAgXIBcgGBAYEBjIGrQatBq0GrQBjAJQAlADFASgBKAEoASgBKAFyAYoBiwG0Ae0B7gHuAlACUQKiAuQC5AMAAxUDFgMuA0cDlQOyA9oESwRLBQIFAwU+BT4FPgVbBVsFawV+BcgFyAXIBcgFyAXIBcgGMQZQBoEG1wdTB4sAegCeAHYAAAAAAAAAAAAAAAAAAACUAJQAlAKBAHMAxQVrA3gCmgEoA0cDLgFyAXICaQGLB1MCHwNNA5UAlAFQAlEBWQBiA7IAzAD3AxwA9wC7AVkAAQatBq0GrQXIBq0FyAUCBQIFAgDeAbwBKAGKAlABigMWAuQG1wEoAe4GBAHtBgQB7QJRACoAlAAAAAAAKgAAAAAAFQB8AHwAAAAAAAIAgAAAA4AFyAADAAcATbgATSsAuAE/RVi4AAAvG7kAAAFCPlm6AAIABgBQK7gAABC4AATcAbgACC+4AAUvuAAIELgAANC4AAAvuAAFELgAA9y4AAAQuAAE3DAxMxEhESUhESGAAwD9gAIA/gAFyPo4gATIAAAAAAEArQFyBnUEVgAGAA+4AE0rALoABAABAFArMDEBESE1IREBBQP7qgRWAXIBcgEolAEo/o4AAAAAAQCtAXIGdQXIAAgAHLgATSsAuAAEL7oAAwAGAFArAboABgADAFArMDETAREhETMRIRGtAXIDwpT7qgLkAXL+2AKa/NL+2AAAAwCM/+cGWAWyABsAMwBHAMi4AE0rALoAIQAVAFArugAHAC8AUCsBuABIL7gAKC+4AEgQuAAA0LgAAC9BBQDaACgA6gAoAAJdQRsACQAoABkAKAApACgAOQAoAEkAKABZACgAaQAoAHkAKACJACgAmQAoAKkAKAC5ACgAyQAoAA1duAAoELgADty4AAAQuAAc3EEbAAYAHAAWABwAJgAcADYAHABGABwAVgAcAGYAHAB2ABwAhgAcAJYAHACmABwAtgAcAMYAHAANXUEFANUAHADlABwAAl0wMRM0PgQzMh4EFRQOBCMiLgQ3FB4CMzI+BDU0LgQjIg4CFzQ+AjMyHgIVFA4CIyIuAow1YIelvmZmvqWIYTU1YYilvmZnvaWHYDV8YafhgFWfiXFRLCxRcYmfVYDhp2GqRnmjXFyjekdHeqNcXKN5RgLMZr6lh2E1NWGHpb5mZ72lh2A1NWCHpb5mgOGnYSxQcIqeVVWfiXFQLGGo4X9co3lHR3mjXFyjekZGeqMAAAACAIz/5wZYBbIAGwAzAMi4AE0rALoAIQAVAFArugAHAC8AUCsBuAA0L7gAKC+4ADQQuAAA0LgAAC9BBQDaACgA6gAoAAJdQRsACQAoABkAKAApACgAOQAoAEkAKABZACgAaQAoAHkAKACJACgAmQAoAKkAKAC5ACgAyQAoAA1duAAoELgADty4AAAQuAAc3EEbAAYAHAAWABwAJgAcADYAHABGABwAVgAcAGYAHAB2ABwAhgAcAJYAHACmABwAtgAcAMYAHAANXUEFANUAHADlABwAAl0wMRM0PgQzMh4EFRQOBCMiLgQ3FB4CMzI+BDU0LgQjIg4CjDVgh6W+Zma+pYhhNTVhiKW+Zme9pYdgNXxhp+GAVZ+JcVEsLFFxiZ9VgOGnYQLMZr6lh2E1NWGHpb5mZ72lh2A1NWCHpb5mgOGnYSxQcIqeVVWfiXFQLGGo4QAAAgClAAAGPwWaAAMABwBNuABNKwC4AT9FWLgAAi8buQACAUI+WboAAQAFAFAruAACELgABNwBuAAIL7gABC+4AAgQuAAA0LgAAC+4AAQQuAAC3LgAABC4AAbcMDETIREhJREhEaUFmvpmBR77XgWa+mZ8BKL7XgAAAwClAAAGPwWaAAMABwAOAGu4AE0rALgBP0VYuAACLxu5AAIBQj5ZugABAAUAUCu4AAIQuAAE3AG4AA8vuAAEL7gADxC4AADQuAAAL7gABBC4AALcuAAAELgABty6AAkAAAACERI5ugALAAAAAhESOboADQAAAAIREjkwMRMhESElESERJQEzFwEzAaUFmvpmBR77XgGW/rz8kAF95P4GBZr6ZnwEovteZgHa5ALn/CMAAAAAAAA8ADwAPAA8AFgAfAFAAeoCJgKAAAAAAAAQAMYAAQAAAAAAAABAAAAAAQAAAAAAAQAFAEAAAQAAAAAAAgAHAEUAAQAAAAAAAwAkAEwAAQAAAAAABAAFAHAAAQAAAAAABQALAHUAAQAAAAAABgAKAIAAAQAAAAAABwAsAIoAAwAABAkAAACAALYAAwAABAkAAQAKATYAAwAABAkAAgAOAUAAAwAABAkAAwBIAU4AAwAABAkABAAKAZYAAwAABAkABQAWAaAAAwAABAkABgAUAbYAAwAABAkABwBYAcpDb3B5cmlnaHQgKGMpIEFzY2Vuc2lvIFN5c3RlbSBTSUEgMjAxMi0yMDE0LiBBbGwgcmlnaHRzIHJlc2VydmVkQVNDVzNSZWd1bGFyVmVyc2lvbiAxLjA7TVM7VmVyc2lvbjEuMDsyMDE0O0ZMNzE0QVNDVzNWZXJzaW9uIDEuMFZlcnNpb24xLjBBU0NXMyBpcyBhIHRyYWRlbWFyayBvZiBBc2NlbnNpbyBTeXN0ZW0gU0lBLgBDAG8AcAB5AHIAaQBnAGgAdAAgACgAYwApACAAQQBzAGMAZQBuAHMAaQBvACAAUwB5AHMAdABlAG0AIABTAEkAQQAgADIAMAAxADIALQAyADAAMQA0AC4AIABBAGwAbAAgAHIAaQBnAGgAdABzACAAcgBlAHMAZQByAHYAZQBkAEEAUwBDAFcAMwBSAGUAZwB1AGwAYQByAFYAZQByAHMAaQBvAG4AIAAxAC4AMAA7AE0AUwA7AFYAZQByAHMAaQBvAG4AMQAuADAAOwAyADAAMQA0ADsARgBMADcAMQA0AEEAUwBDAFcAMwBWAGUAcgBzAGkAbwBuACAAMQAuADAAVgBlAHIAcwBpAG8AbgAxAC4AMABBAFMAQwBXADMAIABpAHMAIABhACAAdAByAGEAZABlAG0AYQByAGsAIABvAGYAIABBAHMAYwBlAG4AcwBpAG8AIABTAHkAcwB0AGUAbQAgAFMASQBBAC4AAgAAAAAAAP8nAJYAAAAAAAAAAAAAAAAAAAAAAAAAAAAKAAABAgEDAQQBBQEGAQcBCAEJAQoETlVMTAd1bmkwMDBEB3VuaUYwMjAHdW5pRjAyMgd1bmlGMDM4B3VuaTI1QzkGY2lyY2xlB3VuaTI2MTAHdW5pMjYxMQAAAAAAAgAQAAX//wAP");
        ascW3.SetStreamIndex(streams_count);
		g_font_files.push(ascW3);

        count_infos = g_font_infos.length;
		g_font_infos[count_infos] = new CFontInfo("ASCW3", 0, g_font_files.length - 1, 0, -1, -1, -1, -1, -1, -1);
		g_map_font_index["ASCW3"] = count_infos;
		/////////////////////////////////////////////////////////////////////

        if (AscFonts.FontPickerByCharacter)
            AscFonts.FontPickerByCharacter.init(window["__fonts_infos"]);

		// удаляем временные переменные
		delete window["__fonts_files"];
		delete window["__fonts_infos"];

        window['AscFonts'].g_font_files = g_font_files;
        window['AscFonts'].g_font_infos = g_font_infos;
	}

    //------------------------------------------------------export------------------------------------------------------
    window['AscFonts'] = window['AscFonts'] || {};
    window['AscFonts'].g_map_font_index = g_map_font_index;
    window['AscFonts'].g_fonts_streams  = g_fonts_streams;

    window['AscFonts'].CFontFileLoader = CFontFileLoader;
    window['AscFonts'].GenerateMapId = GenerateMapId;
    window['AscFonts'].CFontInfo = CFontInfo;
    window['AscFonts'].CFont = CFont;

    window['AscFonts'].CFontFileLoaderEmbed = CFontFileLoaderEmbed;
    window['AscFonts'].CFontInfoEmbed = CFontInfoEmbed;

    window['AscFonts'].ImageLoadStatus = ImageLoadStatus;
    window['AscFonts'].CImage = CImage;

    window['AscFonts'].checkAllFonts = checkAllFonts;

    window['AscFonts'].g_font_infos_embed = [];
    window['AscFonts'].g_map_font_index_embed = {};


    window['AscFonts'].getEmbeddedFontPrefix = function()
    {
        return "Embedded: ";
    };

    window['AscFonts'].initEmbeddedFonts = function(fonts)
    {
        const prefix = AscFonts.getEmbeddedFontPrefix();
        let fontFiles = AscFonts.g_font_files;

        if (CFontFileLoaderEmbed.prototype.startIndex > 0)
            fontFiles.splice(CFontFileLoaderEmbed.prototype.startIndex);

        CFontFileLoaderEmbed.prototype.startIndex = fontFiles.length;
        let currentIndex = CFontFileLoaderEmbed.prototype.startIndex;

        for (let i = 0, len = fonts.length; i < len; i++)
        {
            let name = prefix + fonts[i];
            AscFonts.g_font_files.push(new CFontFileLoaderEmbed(name));
            AscFonts.g_font_infos_embed.push(new CFontInfoEmbed(name, currentIndex++));
            AscFonts.g_map_font_index_embed[name] = i;
        }
    };

    checkAllFonts();

})(window, window.document);

// сначала хотел писать "вытеснение" из этого мапа.
// но тогда нужно хранить base64 строки. Это не круто. По памяти - даже
// выигрыш будет. Не особо то шрифты жмутся lzw или deflate
// поэтому лучше из памяти будем удалять base64 строки
// ----------------------------------------------------------------------------
