/*
 * (c) Copyright Ascensio System SIA 2010-2023
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

    var bIsLocalFontsUse = false;

    function _is_support_cors()
    {
        if (window["NATIVE_EDITOR_ENJINE"] === true)
            return false;

        var xhrSupported = new XMLHttpRequest();
        return !!xhrSupported && ("withCredentials" in xhrSupported);
    }
    var bIsSupportOriginalFormatFonts = _is_support_cors();

    function postLoadScript(scriptName)
    {
        window.postMessage({type: "FROM_PAGE_LOAD_SCRIPT", text: scriptName}, "*");
    }

    /*window.addEventListener("message", function(event) {
        // We only accept messages from ourselves
        if (event.source != window)
            return;

        if (event.data.type && (event.data.type == "FROM_SCRIPT_LOAD_SCRIPT"))
        {
            var _mess = event.data.text;
            var _files = window.g_font_files;

            // потом сделать мап, при первом обращении
            for (var i = 0; i < _files.length; i++)
            {
                if (_files[i].Id == _mess)
                {
                    var bIsUseOrigF = false;
                    if (_files[i].CanUseOriginalFormat && // false if load embedded fonts
                        bIsSupportOriginalFormatFonts) // false if work on ie9
                        bIsUseOrigF = true;

                    if (!bIsUseOrigF)
                    {
                        _files[i]._callback_font_load();
                    }
                    else
                    {
                        bIsLocalFontsUse = false;
                        _files[i].LoadFontAsync(event.data.src, null);
                        bIsLocalFontsUse = true;
                    }

                    break;
                }

            }
        }
        else if (event.data.type && (event.data.type == "FROM_SCRIPT_IS_EXIST"))
        {
            bIsLocalFontsUse = true;
        }
    }, false);*/

    function CFontFileLoader(id)
    {
        this.LoadingCounter = 0;
        this.Id         = id;
        this.Status     = -1;  // -1 - notloaded, 0 - loaded, 1 - error, 2 - loading
        this.stream_index = -1;
        this.callback = null;
        this.IsNeedAddJSToFontPath = true;

        this.CanUseOriginalFormat = true;

        var oThis = this;

        this.CheckLoaded = function()
        {
            return (0 == this.Status || 1 == this.Status);
        };

        this._callback_font_load = function()
        {
            if (!window[oThis.Id])
            {
                oThis.LoadingCounter++;
                if (oThis.LoadingCounter < oThis.GetMaxLoadingCount())
                {
                    //console.log("font loaded: one more attemption");
                    oThis.Status = -1;
                    return;
                }

                oThis.Status = 2; // aka loading...
                var _editor = window["Asc"]["editor"] ? window["Asc"]["editor"] : window.editor;
                _editor.sendEvent("asc_onError", Asc.c_oAscError.ID.CoAuthoringDisconnect, Asc.c_oAscError.Level.Critical);
                return;
            }

            var __font_data_idx = g_fonts_streams.length;
            g_fonts_streams[__font_data_idx] = AscFonts.CreateFontData4(window[oThis.Id]);
            oThis.SetStreamIndex(__font_data_idx);

            oThis.Status = 0;

            // удаляем строку
            delete window[oThis.Id];

            if (null != oThis.callback)
                oThis.callback();
        };

        this.LoadFontAsync2 = function(basePath, _callback)
        {
            this.callback = _callback;
            if (-1 != this.Status)
                return true;

            if (bIsLocalFontsUse)
            {
                postLoadScript(this.Id);
                return;
            }
            this.Status = 2;

            var xhr = new XMLHttpRequest();

            xhr.open('GET', basePath + this.Id, true); // TODO:

            if (typeof ArrayBuffer !== 'undefined' && !window.opera)
                xhr.responseType = 'arraybuffer';

            if (xhr.overrideMimeType)
                xhr.overrideMimeType('text/plain; charset=x-user-defined');
            else
                xhr.setRequestHeader('Accept-Charset', 'x-user-defined');

            xhr.onload = function()
            {
                if (this.status != 200)
                {
                    return this.onerror();
                }

                oThis.Status = 0;

                if (typeof ArrayBuffer !== 'undefined' && !window.opera && this.response)
                {
                    var __font_data_idx = g_fonts_streams.length;
                    var _uintData = new Uint8Array(this.response);
                    g_fonts_streams[__font_data_idx] = new AscFonts.FontStream(_uintData, _uintData.length);
                    oThis.SetStreamIndex(__font_data_idx);
                }
                else if (AscCommon.AscBrowser.isIE)
                {
                    var _response = new VBArray(this["responseBody"]).toArray();

                    var srcLen = _response.length;
                    var stream = new AscFonts.FontStream(AscFonts.allocate(srcLen), srcLen);

                    var dstPx = stream.data;
                    var index = 0;

                    while (index < srcLen)
                    {
                        dstPx[index] = _response[index];
                        index++;
                    }

                    var __font_data_idx = g_fonts_streams.length;
                    g_fonts_streams[__font_data_idx] = stream;
                    oThis.SetStreamIndex(__font_data_idx);
                }
                else
                {
                    var __font_data_idx = g_fonts_streams.length;
                    g_fonts_streams[__font_data_idx] = AscFonts.CreateFontData3(this.responseText);
                    oThis.SetStreamIndex(__font_data_idx);
                }

                // decode
                var guidOdttf = [0xA0, 0x66, 0xD6, 0x20, 0x14, 0x96, 0x47, 0xfa, 0x95, 0x69, 0xB8, 0x50, 0xB0, 0x41, 0x49, 0x48];
                var _stream = g_fonts_streams[g_fonts_streams.length - 1];
                var _data = _stream.data;

                var _count_decode = Math.min(32, _stream.size);
                for (var i = 0; i < _count_decode; ++i)
                    _data[i] ^= guidOdttf[i % 16];

                if (null != oThis.callback)
                    oThis.callback();
            };
            xhr.onerror = function()
            {
                oThis.LoadingCounter++;
                if (oThis.LoadingCounter < oThis.GetMaxLoadingCount())
                {
                    //console.log("font loaded: one more attemption");
                    oThis.Status = -1;
                    return;
                }

                oThis.Status = 2; // aka loading...
                var _editor = window["Asc"]["editor"] ? window["Asc"]["editor"] : window.editor;
                _editor.sendEvent("asc_onError", Asc.c_oAscError.ID.LoadingFontError, Asc.c_oAscError.Level.Critical);
                return;
            };

            xhr.send(null);
        };

        this.LoadFontNative = function()
        {
            var __font_data_idx = g_fonts_streams.length;
            var _data = window["native"]["GetFontBinary"](this.Id);
            g_fonts_streams[__font_data_idx] = new AscFonts.FontStream(_data, _data.length);
            this.SetStreamIndex(__font_data_idx);
            this.Status = 0;
        };
    }

    CFontFileLoader.prototype.GetMaxLoadingCount = function()
    {
        return 3;
    };

    CFontFileLoader.prototype.SetStreamIndex = function(index)
    {
        this.stream_index = index;
    };
    CFontFileLoader.prototype.LoadFontAsync = function(basePath, _callback, isEmbed)
    {
        var oThis = this;
        if (window["AscDesktopEditor"] !== undefined && this.CanUseOriginalFormat)
        {
            if (-1 != this.Status)
                return true;

            this.callback = _callback;
            this.Status = 2;
            window["AscDesktopEditor"]["LoadFontBase64"](this.Id);
            this._callback_font_load();
            return;
        }

        if (this.CanUseOriginalFormat && // false if load embedded fonts
            bIsSupportOriginalFormatFonts) // false if work on ie9
        {
            this.LoadFontAsync2(basePath, _callback);
            return;
        }

        this.callback = _callback;
        if (-1 != this.Status)
            return true;

        this.Status = 2;
        if (bIsLocalFontsUse)
        {
            postLoadScript(this.Id);
            return;
        }

        var scriptElem = document.createElement('script');

        if (scriptElem.readyState && false)
        {
            scriptElem.onreadystatechange = function () {
                if (this.readyState == 'complete' || this.readyState == 'loaded')
                {
                    scriptElem.onreadystatechange = null;
                    setTimeout(oThis._callback_font_load, 0);
                }
            }
        }
        scriptElem.onload = scriptElem.onerror = oThis._callback_font_load;

        var src = basePath + this.Id + ".js";
        if(isEmbed)
            src = AscCommon.g_oDocumentUrls.getUrl(src);
        scriptElem.setAttribute('src', src);
        scriptElem.setAttribute('type','text/javascript');
        document.getElementsByTagName('head')[0].appendChild(scriptElem);
        return false;
    };

    CFontFileLoader.prototype["LoadFontAsync"] = CFontFileLoader.prototype.LoadFontAsync;
    CFontFileLoader.prototype["GetID"] = function() { return this.Id; };
    CFontFileLoader.prototype["GetStatus"] = function() { return this.Status; };
    CFontFileLoader.prototype["GetStreamIndex"] = function() { return this.stream_index; };


    var FONT_TYPE_ADDITIONAL = 0;
    var FONT_TYPE_STANDART = 1;
    var FONT_TYPE_EMBEDDED = 2;
    var FONT_TYPE_ADDITIONAL_CUT = 3;

    var fontstyle_mask_regular = 1;
    var fontstyle_mask_italic = 2;
    var fontstyle_mask_bold = 4;
    var fontstyle_mask_bolditalic = 8;

    function GenerateMapId(api, name, style, size)
    {
        var fontInfo = api.FontLoader.fontInfos[api.FontLoader.map_font_index[name]];
        var index = -1;

        // подбираем шрифт по стилю
        var bNeedBold   = false;
        var bNeedItalic = false;

        var index       = -1;
        var faceIndex   = 0;

        var bSrcItalic  = false;
        var bSrcBold    = false;

        switch (style)
        {
            case FontStyle.FontStyleBoldItalic:
            {
                bSrcItalic  = true;
                bSrcBold    = true;

                bNeedBold   = true;
                bNeedItalic = true;
                if (-1 != fontInfo.indexBI)
                {
                    index = fontInfo.indexBI;
                    faceIndex = fontInfo.faceIndexBI;
                    bNeedBold   = false;
                    bNeedItalic = false;
                }
                else if (-1 != fontInfo.indexB)
                {
                    index = fontInfo.indexB;
                    faceIndex = fontInfo.faceIndexB;
                    bNeedBold = false;
                }
                else if (-1 != fontInfo.indexI)
                {
                    index = fontInfo.indexI;
                    faceIndex = fontInfo.faceIndexI;
                    bNeedItalic = false;
                }
                else
                {
                    index = fontInfo.indexR;
                    faceIndex = fontInfo.faceIndexR;
                }
                break;
            }
            case FontStyle.FontStyleBold:
            {
                bSrcBold    = true;

                bNeedBold   = true;
                bNeedItalic = false;
                if (-1 != fontInfo.indexB)
                {
                    index = fontInfo.indexB;
                    faceIndex = fontInfo.faceIndexB;
                    bNeedBold = false;
                }
                else if (-1 != fontInfo.indexR)
                {
                    index = fontInfo.indexR;
                    faceIndex = fontInfo.faceIndexR;
                }
                else if (-1 != fontInfo.indexBI)
                {
                    index = fontInfo.indexBI;
                    faceIndex = fontInfo.faceIndexBI;
                    bNeedBold = false;
                }
                else
                {
                    index = fontInfo.indexI;
                    faceIndex = fontInfo.faceIndexI;
                }
                break;
            }
            case FontStyle.FontStyleItalic:
            {
                bSrcItalic  = true;

                bNeedBold   = false;
                bNeedItalic = true;
                if (-1 != fontInfo.indexI)
                {
                    index = fontInfo.indexI;
                    faceIndex = fontInfo.faceIndexI;
                    bNeedItalic = false;
                }
                else if (-1 != fontInfo.indexR)
                {
                    index = fontInfo.indexR;
                    faceIndex = fontInfo.faceIndexR;
                }
                else if (-1 != fontInfo.indexBI)
                {
                    index = fontInfo.indexBI;
                    faceIndex = fontInfo.faceIndexBI;
                    bNeedItalic = false;
                }
                else
                {
                    index = fontInfo.indexB;
                    faceIndex = fontInfo.faceIndexB;
                }
                break;
            }
            case FontStyle.FontStyleRegular:
            {
                bNeedBold   = false;
                bNeedItalic = false;
                if (-1 != fontInfo.indexR)
                {
                    index = fontInfo.indexR;
                    faceIndex = fontInfo.faceIndexR;
                }
                else if (-1 != fontInfo.indexI)
                {
                    index = fontInfo.indexI;
                    faceIndex = fontInfo.faceIndexI;
                }
                else if (-1 != fontInfo.indexB)
                {
                    index = fontInfo.indexB;
                    faceIndex = fontInfo.faceIndexB;
                }
                else
                {
                    index = fontInfo.indexBI;
                    faceIndex = fontInfo.faceIndexBI;
                }
            }
        }

        var _ext = "";
        if (bNeedBold)
            _ext += "nbold";
        if (bNeedItalic)
            _ext += "nitalic";

        // index != -1 (!!!)
        var fontfile = api.FontLoader.fontFiles[index];
        return fontfile.Id + faceIndex + size + _ext;
    }

    function CFontInfo(sName, thumbnail, type, indexR, faceIndexR, indexI, faceIndexI, indexB, faceIndexB, indexBI, faceIndexBI)
    {
        this.Name = sName;
        this.Thumbnail = thumbnail;
        this.Type = type;
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
        CheckFontLoadStyles : function(global_loader)
        {
            if ((this.NeedStyles & 0x0F) == 0x0F)
            {
                this.needR = true;
                this.needI = true;
                this.needB = true;
                this.needBI = true;
            }
            else
            {
                if ((this.NeedStyles & 1) != 0)
                {
                    // нужен стиль regular
                    if (false === this.needR)
                    {
                        this.needR = true;
                        if (-1 == this.indexR)
                        {
                            if (-1 != this.indexI)
                            {
                                this.needI = true;
                            }
                            else if (-1 != this.indexB)
                            {
                                this.needB = true;
                            }
                            else
                            {
                                this.needBI = true;
                            }
                        }
                    }
                }
                if ((this.NeedStyles & 2) != 0)
                {
                    // нужен стиль italic
                    if (false === this.needI)
                    {
                        this.needI = true;
                        if (-1 == this.indexI)
                        {
                            if (-1 != this.indexR)
                            {
                                this.needR = true;
                            }
                            else if (-1 != this.indexBI)
                            {
                                this.needBI = true;
                            }
                            else
                            {
                                this.needB = true;
                            }
                        }
                    }
                }
                if ((this.NeedStyles & 4) != 0)
                {
                    // нужен стиль bold
                    if (false === this.needB)
                    {
                        this.needB = true;
                        if (-1 == this.indexB)
                        {
                            if (-1 != this.indexR)
                            {
                                this.needR = true;
                            }
                            else if (-1 != this.indexBI)
                            {
                                this.needBI = true;
                            }
                            else
                            {
                                this.needI = true;
                            }
                        }
                    }
                }
                if ((this.NeedStyles & 8) != 0)
                {
                    // нужен стиль bold
                    if (false === this.needBI)
                    {
                        this.needBI = true;
                        if (-1 == this.indexBI)
                        {
                            if (-1 != this.indexB)
                            {
                                this.needB = true;
                            }
                            else if (-1 != this.indexI)
                            {
                                this.needI = true;
                            }
                            else
                            {
                                this.needR = true;
                            }
                        }
                    }
                }
            }

            var isEmbed = (FONT_TYPE_EMBEDDED == this.Type);
            var fonts = isEmbed ? global_loader.embeddedFontFiles : global_loader.fontFiles;
            var basePath = isEmbed ? global_loader.embeddedFilesPath : global_loader.fontFilesPath;
            var isNeed = false;
            if ((this.needR === true) && (-1 != this.indexR) && (fonts[this.indexR].CheckLoaded() === false))
            {
                fonts[this.indexR].LoadFontAsync(basePath, null, isEmbed);
                isNeed = true;
            }
            if ((this.needI === true) && (-1 != this.indexI) && (fonts[this.indexI].CheckLoaded() === false))
            {
                fonts[this.indexI].LoadFontAsync(basePath, null, isEmbed);
                isNeed = true;
            }
            if ((this.needB === true) && (-1 != this.indexB) && (fonts[this.indexB].CheckLoaded() === false))
            {
                fonts[this.indexB].LoadFontAsync(basePath, null, isEmbed);
                isNeed = true;
            }
            if ((this.needBI === true) && (-1 != this.indexBI) && (fonts[this.indexBI].CheckLoaded() === false))
            {
                fonts[this.indexBI].LoadFontAsync(basePath, null, isEmbed);
                isNeed = true;
            }

            return isNeed;
        },

        CheckFontLoadStylesNoLoad : function(global_loader)
        {
            var fonts = (FONT_TYPE_EMBEDDED == this.Type) ? global_loader.embeddedFontFiles : global_loader.fontFiles;
            var _isNeed = false;
            if ((-1 != this.indexR) && (fonts[this.indexR].CheckLoaded() === false))
            {
                _isNeed = true;
            }
            if ((-1 != this.indexI) && (fonts[this.indexI].CheckLoaded() === false))
            {
                _isNeed = true;
            }
            if ((-1 != this.indexB) && (fonts[this.indexB].CheckLoaded() === false))
            {
                _isNeed = true;
            }
            if ((-1 != this.indexBI) && (fonts[this.indexBI].CheckLoaded() === false))
            {
                _isNeed = true;
            }

            return _isNeed;
        },

        LoadFontsFromServer : function(global_loader)
        {
            var fonts = global_loader.fontFiles;
            var basePath = global_loader.fontFilesPath;
            if ((-1 != this.indexR) && (fonts[this.indexR].CheckLoaded() === false))
            {
                fonts[this.indexR].LoadFontAsync(basePath, null);
            }
            if ((-1 != this.indexI) && (fonts[this.indexI].CheckLoaded() === false))
            {
                fonts[this.indexI].LoadFontAsync(basePath, null);
            }
            if ((-1 != this.indexB) && (fonts[this.indexB].CheckLoaded() === false))
            {
                fonts[this.indexB].LoadFontAsync(basePath, null);
            }
            if ((-1 != this.indexBI) && (fonts[this.indexBI].CheckLoaded() === false))
            {
                fonts[this.indexBI].LoadFontAsync(basePath, null);
            }
        },

        LoadFont : function(font_loader, fontManager, fEmSize, lStyle, dHorDpi, dVerDpi, transform, isNoSetupToManager)
        {
            // подбираем шрифт по стилю
            var sReturnName = this.Name;
            var bNeedBold   = false;
            var bNeedItalic = false;

            var index       = -1;
            var faceIndex   = 0;

            var bSrcItalic  = false;
            var bSrcBold    = false;

            switch (lStyle)
            {
                case FontStyle.FontStyleBoldItalic:
                {
                    bSrcItalic  = true;
                    bSrcBold    = true;

                    bNeedBold   = true;
                    bNeedItalic = true;
                    if (-1 != this.indexBI)
                    {
                        index = this.indexBI;
                        faceIndex = this.faceIndexBI;
                        bNeedBold   = false;
                        bNeedItalic = false;
                    }
                    else if (-1 != this.indexB)
                    {
                        index = this.indexB;
                        faceIndex = this.faceIndexB;
                        bNeedBold = false;
                    }
                    else if (-1 != this.indexI)
                    {
                        index = this.indexI;
                        faceIndex = this.faceIndexI;
                        bNeedItalic = false;
                    }
                    else
                    {
                        index = this.indexR;
                        faceIndex = this.faceIndexR;
                    }
                    break;
                }
                case FontStyle.FontStyleBold:
                {
                    bSrcBold    = true;

                    bNeedBold   = true;
                    bNeedItalic = false;
                    if (-1 != this.indexB)
                    {
                        index = this.indexB;
                        faceIndex = this.faceIndexB;
                        bNeedBold = false;
                    }
                    else if (-1 != this.indexR)
                    {
                        index = this.indexR;
                        faceIndex = this.faceIndexR;
                    }
                    else if (-1 != this.indexBI)
                    {
                        index = this.indexBI;
                        faceIndex = this.faceIndexBI;
                        bNeedBold = false;
                    }
                    else
                    {
                        index = this.indexI;
                        faceIndex = this.faceIndexI;
                    }
                    break;
                }
                case FontStyle.FontStyleItalic:
                {
                    bSrcItalic  = true;

                    bNeedBold   = false;
                    bNeedItalic = true;
                    if (-1 != this.indexI)
                    {
                        index = this.indexI;
                        faceIndex = this.faceIndexI;
                        bNeedItalic = false;
                    }
                    else if (-1 != this.indexR)
                    {
                        index = this.indexR;
                        faceIndex = this.faceIndexR;
                    }
                    else if (-1 != this.indexBI)
                    {
                        index = this.indexBI;
                        faceIndex = this.faceIndexBI;
                        bNeedItalic = false;
                    }
                    else
                    {
                        index = this.indexB;
                        faceIndex = this.faceIndexB;
                    }
                    break;
                }
                case FontStyle.FontStyleRegular:
                {
                    bNeedBold   = false;
                    bNeedItalic = false;
                    if (-1 != this.indexR)
                    {
                        index = this.indexR;
                        faceIndex = this.faceIndexR;
                    }
                    else if (-1 != this.indexI)
                    {
                        index = this.indexI;
                        faceIndex = this.faceIndexI;
                    }
                    else if (-1 != this.indexB)
                    {
                        index = this.indexB;
                        faceIndex = this.faceIndexB;
                    }
                    else
                    {
                        index = this.indexBI;
                        faceIndex = this.faceIndexBI;
                    }
                }
            }

            // index != -1 (!!!)
            var fontfile = null;
            if (this.Type == FONT_TYPE_EMBEDDED)
                fontfile = font_loader.embeddedFontFiles[index];
            else
                fontfile = font_loader.fontFiles[index];

            if (window["NATIVE_EDITOR_ENJINE"] && fontfile.Status != 0)
            {
                fontfile.LoadFontNative();
            }

            var pFontFile = fontManager.LoadFont(fontfile, faceIndex, fEmSize, bSrcBold, bSrcItalic, bNeedBold, bNeedItalic, isNoSetupToManager);

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

        GetFontID : function(font_loader, lStyle)
        {
            // подбираем шрифт по стилю
            var sReturnName = this.Name;
            var bNeedBold   = false;
            var bNeedItalic = false;

            var index       = -1;
            var faceIndex   = 0;

            var bSrcItalic  = false;
            var bSrcBold    = false;

            switch (lStyle)
            {
                case FontStyle.FontStyleBoldItalic:
                {
                    bSrcItalic  = true;
                    bSrcBold    = true;

                    bNeedBold   = true;
                    bNeedItalic = true;
                    if (-1 != this.indexBI)
                    {
                        index = this.indexBI;
                        faceIndex = this.faceIndexBI;
                        bNeedBold   = false;
                        bNeedItalic = false;
                    }
                    else if (-1 != this.indexB)
                    {
                        index = this.indexB;
                        faceIndex = this.faceIndexB;
                        bNeedBold = false;
                    }
                    else if (-1 != this.indexI)
                    {
                        index = this.indexI;
                        faceIndex = this.faceIndexI;
                        bNeedItalic = false;
                    }
                    else
                    {
                        index = this.indexR;
                        faceIndex = this.faceIndexR;
                    }
                    break;
                }
                case FontStyle.FontStyleBold:
                {
                    bSrcBold    = true;

                    bNeedBold   = true;
                    bNeedItalic = false;
                    if (-1 != this.indexB)
                    {
                        index = this.indexB;
                        faceIndex = this.faceIndexB;
                        bNeedBold = false;
                    }
                    else if (-1 != this.indexR)
                    {
                        index = this.indexR;
                        faceIndex = this.faceIndexR;
                    }
                    else if (-1 != this.indexBI)
                    {
                        index = this.indexBI;
                        faceIndex = this.faceIndexBI;
                        bNeedBold = false;
                    }
                    else
                    {
                        index = this.indexI;
                        faceIndex = this.faceIndexI;
                    }
                    break;
                }
                case FontStyle.FontStyleItalic:
                {
                    bSrcItalic  = true;

                    bNeedBold   = false;
                    bNeedItalic = true;
                    if (-1 != this.indexI)
                    {
                        index = this.indexI;
                        faceIndex = this.faceIndexI;
                        bNeedItalic = false;
                    }
                    else if (-1 != this.indexR)
                    {
                        index = this.indexR;
                        faceIndex = this.faceIndexR;
                    }
                    else if (-1 != this.indexBI)
                    {
                        index = this.indexBI;
                        faceIndex = this.faceIndexBI;
                        bNeedItalic = false;
                    }
                    else
                    {
                        index = this.indexB;
                        faceIndex = this.faceIndexB;
                    }
                    break;
                }
                case FontStyle.FontStyleRegular:
                {
                    bNeedBold   = false;
                    bNeedItalic = false;
                    if (-1 != this.indexR)
                    {
                        index = this.indexR;
                        faceIndex = this.faceIndexR;
                    }
                    else if (-1 != this.indexI)
                    {
                        index = this.indexI;
                        faceIndex = this.faceIndexI;
                    }
                    else if (-1 != this.indexB)
                    {
                        index = this.indexB;
                        faceIndex = this.faceIndexB;
                    }
                    else
                    {
                        index = this.indexBI;
                        faceIndex = this.faceIndexBI;
                    }
                }
            }

            // index != -1 (!!!)
            var fontfile = (this.Type == FONT_TYPE_EMBEDDED) ? font_loader.embeddedFontFiles[index] : font_loader.fontFiles[index];
            return { id: fontfile.Id, faceIndex : faceIndex, file : fontfile };
        },

        GetBaseStyle : function(lStyle)
        {
            switch (lStyle)
            {
                case FontStyle.FontStyleBoldItalic:
                {
                    if (-1 != this.indexBI)
                    {
                        return FontStyle.FontStyleBoldItalic;
                    }
                    else if (-1 != this.indexB)
                    {
                        return FontStyle.FontStyleBold;
                    }
                    else if (-1 != this.indexI)
                    {
                        return FontStyle.FontStyleItalic;
                    }
                    else
                    {
                        return FontStyle.FontStyleRegular;
                    }
                    break;
                }
                case FontStyle.FontStyleBold:
                {
                    if (-1 != this.indexB)
                    {
                        return FontStyle.FontStyleBold;
                    }
                    else if (-1 != this.indexR)
                    {
                        return FontStyle.FontStyleRegular;
                    }
                    else if (-1 != this.indexBI)
                    {
                        return FontStyle.FontStyleBoldItalic;
                    }
                    else
                    {
                        return FontStyle.FontStyleItalic;
                    }
                    break;
                }
                case FontStyle.FontStyleItalic:
                {
                    if (-1 != this.indexI)
                    {
                        return FontStyle.FontStyleItalic;
                    }
                    else if (-1 != this.indexR)
                    {
                        return FontStyle.FontStyleRegular;
                    }
                    else if (-1 != this.indexBI)
                    {
                        return FontStyle.FontStyleBoldItalic;
                    }
                    else
                    {
                        return FontStyle.FontStyleBold;
                    }
                    break;
                }
                case FontStyle.FontStyleRegular:
                {
                    if (-1 != this.indexR)
                    {
                        return FontStyle.FontStyleRegular;
                    }
                    else if (-1 != this.indexI)
                    {
                        return FontStyle.FontStyleItalic;
                    }
                    else if (-1 != this.indexB)
                    {
                        return FontStyle.FontStyleBold;
                    }
                    else
                    {
                        return FontStyle.FontStyleBoldItalic;
                    }
                }
            }
            return FontStyle.FontStyleRegular;
        }
    };

    // здесь если type == FONT_TYPE_EMBEDDED, то thumbnail - это base64 картинка,
    // иначе - это позиция (y) в общем табнейле всех шрифтов (ADDITIONAL и STANDART)
    function CFont(name, id, type, thumbnail, style)
    {
        this.name = name;
        this.id = id;
        this.type = type;
        this.thumbnail = thumbnail;
        if(null != style)
            this.NeedStyles = style;
        else
            this.NeedStyles = fontstyle_mask_regular | fontstyle_mask_italic | fontstyle_mask_bold | fontstyle_mask_bolditalic;
    }
    CFont.prototype.asc_getFontId = function() { return this.id; };
    CFont.prototype.asc_getFontName = function()
    {
        var _name = AscFonts.g_fontApplication ? AscFonts.g_fontApplication.NameToInterface[this.name] : null;
        return _name ? _name : this.name;
    };
    CFont.prototype.asc_getFontThumbnail = function() { return this.thumbnail; };
    CFont.prototype.asc_getFontType = function() { return this.type; };

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

	var g_font_files, g_font_infos;
	function checkAllFonts()
    {
        var i, l;
        var files = window["__fonts_files"];
		if (!files && window["native"] && window["native"]["GenerateAllFonts"])
		{
			// тогда должны быть глобальные переменные такие, без window
			window["native"]["GenerateAllFonts"]();
		}

		files = window["__fonts_files"];
		l = files ? files.length : 0;
		g_font_files = new Array(l);
		for (i = 0; i < l; i++)
		{
			g_font_files[i] = new CFontFileLoader(files[i]);
		}

		files = window["__fonts_infos"];
		l = files ? files.length : 0;
		g_font_infos = new Array(l);

		var curIndex = 0;
		var ascFontFath = "ASC.ttf";
		for (i = 0; i < l; i++)
		{
			var _info = files[i];
			if ("ASCW3" === _info[0])
			{
			    ascFontFath = g_font_files[_info[1]].Id;
                continue;
            }

			g_font_infos[curIndex] = new CFontInfo(_info[0], i, 0, _info[1], _info[2], _info[3], _info[4], _info[5], _info[6], _info[7], _info[8]);
			g_map_font_index[_info[0]] = curIndex;
			curIndex++;
		}

        g_font_infos.length = curIndex;

		/////////////////////////////////////////////////////////////////////
		// а вот это наш шрифт - аналог wingdings3
		var _wngds3 = new CFontFileLoader(ascFontFath);
		_wngds3.Status = 0;
		l = g_fonts_streams.length;
		g_fonts_streams[l] = AscFonts.CreateFontData2("AAEAAAARAQAABAAQTFRTSOGXTFYAAAIgAAAADk9TLzJ0/uscAAABmAAAAGBWRE1Yblp14wAAAjAAAAXgY21hcNCBcnoAAAksAAAAaGN2dCBBCTniAAAUWAAAAohmcGdtu26+2AAACZQAAAi4Z2FzcAAPABYAAB9cAAAADGdseWaWCCxJAAAW4AAABQBoZG1466dcHQAACBAAAAEcaGVhZBMko9YAAAEcAAAANmhoZWEOeQN0AAABVAAAACRobXR4MfIEPAAAAfgAAAAobG9jYQVeBDYAABvgAAAAFm1heHAIXQkWAAABeAAAACBuYW1lZvQXkQAAG/gAAALocG9zdKnSYBIAAB7gAAAAenByZXDMvTRHAAASTAAAAgsAAQAAAAEAAFPp5kFfDzz1ABsIAAAAAADPTriwAAAAAN2yoHYAgP/nBnUFyAAAAAwAAQAAAAAAAAABAAAHPv5OAEMHIQCA/foGdQABAAAAAAAAAAAAAAAAAAAACgABAAAACgBIAAMAAAAAAAIAEAAUAFcAAAfoCLgAAAAAAAMFjAGQAAUAAATOBM4AAAMWBM4EzgAAAxYAZgISDAAFBAECAQgHBwcHAAAAAAAAAAAAAAAAAAAAAE1TICAAQCXJ8DgF0/5RATMHPgGygAAAAAAAAAAEJgW7AAAAIAAABAAAgAAAAAAE0gAAAgAAAAchAK0EbwCtBuQAjAbkAIwG5AClBuQApQAAAApLAQEBS0tLS0tLAAAAAAABAAEBAQEBAAwA+Aj/AAgACP/+AAkACf/+AAoACv/9AAsACv/9AAwAC//9AA0ADP/9AA4ADf/9AA8ADv/8ABAAD//8ABEAEP/8ABIAEf/8ABMAEv/7ABQAE//7ABUAFP/7ABYAFP/7ABcAFf/7ABgAFv/6ABkAF//6ABoAGP/6ABsAGf/6ABwAGv/6AB0AG//5AB4AHP/5AB8AHf/5ACAAHf/5ACEAHv/5ACIAH//4ACMAIP/4ACQAIf/4ACUAIv/4ACYAI//3ACcAJP/3ACgAJf/3ACkAJv/3ACoAJ//3ACsAJ//2ACwAKP/2AC0AKf/2AC4AKv/2AC8AK//2ADAALP/1ADEALf/1ADIALv/1ADMAL//1ADQAMP/0ADUAMP/0ADYAMf/0ADcAMv/0ADgAM//0ADkANP/zADoANf/zADsANv/zADwAN//zAD0AOP/zAD4AOf/yAD8AOv/yAEAAOv/yAEEAO//yAEIAPP/yAEMAPf/xAEQAPv/xAEUAP//xAEYAQP/xAEcAQf/wAEgAQv/wAEkAQ//wAEoAQ//wAEsARP/wAEwARf/vAE0ARv/vAE4AR//vAE8ASP/vAFAASf/vAFEASv/uAFIAS//uAFMATP/uAFQATf/uAFUATf/tAFYATv/tAFcAT//tAFgAUP/tAFkAUf/tAFoAUv/sAFsAU//sAFwAVP/sAF0AVf/sAF4AVv/sAF8AV//rAGAAV//rAGEAWP/rAGIAWf/rAGMAWv/rAGQAW//qAGUAXP/qAGYAXf/qAGcAXv/qAGgAX//pAGkAYP/pAGoAYP/pAGsAYf/pAGwAYv/pAG0AY//oAG4AZP/oAG8AZf/oAHAAZv/oAHEAZ//oAHIAaP/nAHMAaf/nAHQAav/nAHUAav/nAHYAa//mAHcAbP/mAHgAbf/mAHkAbv/mAHoAb//mAHsAcP/lAHwAcf/lAH0Acv/lAH4Ac//lAH8Ac//lAIAAdP/kAIEAdf/kAIIAdv/kAIMAd//kAIQAeP/kAIUAef/jAIYAev/jAIcAe//jAIgAfP/jAIkAff/iAIoAff/iAIsAfv/iAIwAf//iAI0AgP/iAI4Agf/hAI8Agv/hAJAAg//hAJEAhP/hAJIAhf/hAJMAhv/gAJQAhv/gAJUAh//gAJYAiP/gAJcAif/gAJgAiv/fAJkAi//fAJoAjP/fAJsAjf/fAJwAjv/eAJ0Aj//eAJ4AkP/eAJ8AkP/eAKAAkf/eAKEAkv/dAKIAk//dAKMAlP/dAKQAlf/dAKUAlv/dAKYAl//cAKcAmP/cAKgAmf/cAKkAmf/cAKoAmv/bAKsAm//bAKwAnP/bAK0Anf/bAK4Anv/bAK8An//aALAAoP/aALEAof/aALIAov/aALMAo//aALQAo//ZALUApP/ZALYApf/ZALcApv/ZALgAp//ZALkAqP/YALoAqf/YALsAqv/YALwAq//YAL0ArP/XAL4Arf/XAL8Arf/XAMAArv/XAMEAr//XAMIAsP/WAMMAsf/WAMQAsv/WAMUAs//WAMYAtP/WAMcAtf/VAMgAtv/VAMkAtv/VAMoAt//VAMsAuP/UAMwAuf/UAM0Auv/UAM4Au//UAM8AvP/UANAAvf/TANEAvv/TANIAv//TANMAwP/T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");
		_wngds3.SetStreamIndex(l);
		g_font_files.push(_wngds3);

		l = g_font_infos.length;
		g_font_infos[l] = new CFontInfo("ASCW3", 0, FONT_TYPE_ADDITIONAL, g_font_files.length - 1, 0, -1, -1, -1, -1, -1, -1);
		g_map_font_index["ASCW3"] = l;
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
    window['AscFonts'].g_fonts_streams = g_fonts_streams;

    window['AscFonts'].FONT_TYPE_ADDITIONAL = FONT_TYPE_ADDITIONAL;
    window['AscFonts'].FONT_TYPE_STANDART = FONT_TYPE_STANDART;
    window['AscFonts'].FONT_TYPE_EMBEDDED = FONT_TYPE_EMBEDDED;
    window['AscFonts'].FONT_TYPE_ADDITIONAL_CUT = FONT_TYPE_ADDITIONAL_CUT;

    window['AscFonts'].CFontFileLoader = CFontFileLoader;
    window['AscFonts'].GenerateMapId = GenerateMapId;
    window['AscFonts'].CFontInfo = CFontInfo;
    window['AscFonts'].CFont = CFont;

    var prot = CFont.prototype;
    prot['asc_getFontId'] = prot.asc_getFontId;
    prot['asc_getFontName'] = prot.asc_getFontName;
    prot['asc_getFontThumbnail'] = prot.asc_getFontThumbnail;
    prot['asc_getFontType'] = prot.asc_getFontType;
    window['AscFonts'].ImageLoadStatus = ImageLoadStatus;

    window['AscFonts'].CImage = CImage;

    window['AscFonts'].checkAllFonts = checkAllFonts;

    checkAllFonts();

})(window, window.document);

// сначала хотел писать "вытеснение" из этого мапа.
// но тогда нужно хранить base64 строки. Это не круто. По памяти - даже
// выигрыш будет. Не особо то шрифты жмутся lzw или deflate
// поэтому лучше из памяти будем удалять base64 строки
// ----------------------------------------------------------------------------
