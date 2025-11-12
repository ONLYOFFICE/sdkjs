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

window["AscNotLoadAllScript"] = true;

const fontslot_None     = 0x00;
const fontslot_ASCII    = 0x01;
const fontslot_EastAsia = 0x02;
const fontslot_CS       = 0x04;
const fontslot_HAnsi    = 0x08;
const fontslot_Unknown  = 0x10;

window.editor = {};
window.editor.asyncFontsDocumentStartLoaded = function() {};
window.editor.asyncFontsDocumentEndLoaded = function() {

    doDraw();
};

window['AscCommon'] = window['AscCommon'] || {};
window["Asc"] = window["Asc"] || {};

AscCommon.loadScript = function(url, onSuccess, onError)
{
    var script = document.createElement('script');
    script.type = 'text/javascript';
    script.src = url;
    script.onload = onSuccess;
    script.onerror = onError;

    // Fire the loading
    document.head.appendChild(script);
};

window.onLoadFontsModuleOld = window.onLoadFontsModule;
window.onLoadFontsModule = function(window, undefined) 
{
    onLoadFontsModuleOld(window, undefined);

    AscFonts.g_fontManagerMeasurer = new AscFonts.CFontManager();
    AscFonts.g_fontManagerMeasurer.Initialize();

    AscFonts.g_fontManagerRenderer = new AscFonts.CFontManager();
    AscFonts.g_fontManagerRenderer.Initialize();
    AscFonts.g_fontManagerRenderer.SetHintsProps(true, true);
}

AscCommon.isLeadingSurrogateChar = function(nCharCode)
{
    return (nCharCode >= 0xD800 && nCharCode <= 0xDFFF);
};
AscCommon.decodeSurrogateChar = function(nLeadingChar, nTrailingChar)
{
    if (nLeadingChar < 0xDC00 && nTrailingChar >= 0xDC00 && nTrailingChar <= 0xDFFF)
        return 0x10000 + ((nLeadingChar & 0x3FF) << 10) | (nTrailingChar & 0x3FF);
    else
        return null;
};
AscCommon.encodeSurrogateChar = function(nUnicode)
{
    if (nUnicode < 0x10000)
    {
        return String.fromCharCode(nUnicode);
    }
    else
    {
        nUnicode = nUnicode - 0x10000;
        var nLeadingChar = 0xD800 | (nUnicode >> 10);
        var nTrailingChar = 0xDC00 | (nUnicode & 0x3FF);
        return String.fromCharCode(nLeadingChar) + String.fromCharCode(nTrailingChar);
    }
};

function CUnicodeIterator(str)
{
    this._position = 0;
    this._index = 0;
    this._str = str;
}
CUnicodeIterator.prototype =
{
    isOutside : function()
    {
        return (this._index >= this._str.length);
    },
    isInside : function()
    {
        return (this._index < this._str.length);
    },
    value : function()
    {
        if (this._index >= this._str.length)
            return 0;

        var nCharCode = this._str.charCodeAt(this._index);
        if (!AscCommon.isLeadingSurrogateChar(nCharCode))
            return nCharCode;

        if ((this._str.length - 1) == this._index)
            return nCharCode; // error

        var nTrailingChar = this._str.charCodeAt(this._index + 1);
        return AscCommon.decodeSurrogateChar(nCharCode, nTrailingChar);
    },
    next : function()
    {
        if (this._index >= this._str.length)
            return;

        this._position++;
        if (!AscCommon.isLeadingSurrogateChar(this._str.charCodeAt(this._index)))
        {
            ++this._index;
            return;
        }

        if (this._index == (this._str.length - 1))
        {
            ++this._index;
            return;
        }

        this._index += 2;
    },
    position : function()
    {
        return this._position;
    }
};

CUnicodeIterator.prototype.check = CUnicodeIterator.prototype.isInside;
String.prototype.getUnicodeIterator = function()
{
    return new CUnicodeIterator(this);
};

function getScriptByText()
{
    let scripts = {};
    let text = document.getElementById("id_textarea").value;
    for (let iter = text.getUnicodeIterator(); iter.check(); iter.next())
    {
        let script = AscFonts.hb_get_script_by_unicode(iter.value());
        if (!scripts["" + script]) scripts["" + script] = 0;
        scripts["" + script]++;
    }

    let currentScript = AscFonts.HB_SCRIPT.HB_SCRIPT_COMMON;
    let currentScriptMax = 0;
    for (let item in scripts)
    {
        if (scripts[item] > currentScriptMax)
        {
            currentScript = parseInt(item);
            currentScriptMax = scripts[item];
        }
    }

    return currentScript;
}

function doDraw()
{
    let _font = AscFonts.g_fontApplication.GetFontFileWeb(document.getElementById("id_fonts").value, 0);
    let font_name_index = AscFonts.g_map_font_index[_font.m_wsFontName];

    let fontSize = 40;

    let fontFile = AscFonts.g_font_infos[font_name_index].LoadFont(AscCommon.g_font_loader, AscFonts.g_fontManagerMeasurer, fontSize, 0, 72, 72, undefined);

    // FEATURES
    var feature = 0;
    if (document.getElementById("id_lig_standard").checked)
        feature |= AscFonts.HB_FEATURE.HB_FEATURE_TYPE_LIGATURES_STANDARD;
    if (document.getElementById("id_lig_cont").checked)
        feature |= AscFonts.HB_FEATURE.HB_FEATURE_TYPE_LIGATURES_CONTEXTUAL;
    if (document.getElementById("id_lig_hist").checked)
        feature |= AscFonts.HB_FEATURE.HB_FEATURE_TYPE_LIGATURES_HISTORICAL;
    if (document.getElementById("id_lig_disc").checked)
        feature |= AscFonts.HB_FEATURE.HB_FEATURE_TYPE_LIGATURES_DISCRETIONARY;

    if (document.getElementById("id_kerning").checked)
        feature |= AscFonts.HB_FEATURE.HB_FEATURE_TYPE_KERNING;

    // SCRIPT
    let script_string = document.getElementById("id_script").value;
    script_string = script_string.toUpperCase();
    let script = AscFonts.HB_SCRIPT["HB_SCRIPT_" + script_string];

    // DIRECTION
    let direction = AscFonts.HB_DIRECTION.HB_DIRECTION_LTR;
    if ("RTL" === document.getElementById("id_direction").value)
        direction = AscFonts.HB_DIRECTION.HB_DIRECTION_RTL;

    let segments = fontFile.ShapeText(document.getElementById("id_textarea").value, feature, script, direction,"en");

    let mapType = [
        "UNCLASSIFIED",
        "BASE_GLYPH",
        "LIGATURE",
        "MARK",
        "COMPONENT"
    ];
    let mapFlags = [
        "NONE",
        "UNSAFE_TO_BREAK",
        "UNSAFE_TO_CONCAT",
        "DEFINED"
    ];

    let logSegments = [];
    for (let i = 0, len = segments.length; i < len; i++)
    {
        let seg = {
            font: segments[i].font,
            text: segments[i].text,
            glyphs: []
        };

        for (let j = 0, glyphCount = segments[i].glyphs.length; j < glyphCount; j++)
        {
            let input = segments[i].glyphs[j];
            let output = {};

            output.type = mapType[input.type];
            output.flags = mapFlags[input.flags];

            output.gid = input.gid;
            output.cluster = input.cluster;

            output.text = input.text;

            output.x_advance = input.x_advance;
            output.y_advance = input.y_advance;
            output.x_offset = input.x_offset;
            output.y_offset = input.y_offset;

            seg.glyphs.push(output);
        }

        logSegments.push(seg);
    }

    //if (direction === AscFonts.HB_DIRECTION.HB_DIRECTION_RTL)
    //    logSegments = logSegments.reverse();

    let canvas = document.getElementById("id_canvas");
    let canvasCtx = canvas.getContext('2d');
    canvas.width = canvas.clientWidth;
    canvas.height = canvas.clientHeight;

    let wPx = canvas.width;
    let hPx = canvas.height;
    let wMm = 25.4 * wPx / 96;
    let hMm = 25.4 * hPx / 96;

    let g = new AscCommon.CGraphics();
    g.init(canvasCtx, wPx, hPx, wMm, hMm);
    g.m_oFontManager = AscFonts.g_fontManagerRenderer;
    g.b_color1(0, 0, 0, 255);

    let margin = 5;
    let positionX = margin;
    //if (direction === AscFonts.HB_DIRECTION.HB_DIRECTION_RTL)
    //    positionX = wPx - positionX - 100;
    let positionY = 3 * hPx / 4;

    let currentX = 0;
    let currentY = 0;

    function coordToGraphics(coord) {
        return (coord >> 6) * 96 / 72;
    }

    for (let i = 0, len = logSegments.length; i < len; i++)
    {
        g.m_oTextPr = {
            RFonts : { Ascii : { Name : logSegments[i].font, Index : -1 } },
            FontSize : (((fontSize * 2) + 0.5) >> 0) / 2,
            Bold : false,
            Italic : false
        };
        g.m_oGrFonts.Ascii.Name = g.m_oTextPr.RFonts.Ascii.Name;
        g.m_oGrFonts.Ascii.Index = -1;
        g.SetFontSlot(fontslot_ASCII, 1);

        for (let j = 0, glyphsCount = logSegments[i].glyphs.length; j < glyphsCount; j++)
        {
            let glyph = segments[i].glyphs[j];

            let drawX = currentX + glyph.x_offset;
            let drawY = currentY - glyph.y_offset;

            currentX += glyph.x_advance;
            //currentY += glyph.y_advance;

            drawX = positionX + coordToGraphics(drawX);
            drawY = positionY + coordToGraphics(drawY);

            g.tg(glyph.gid, drawX, drawY);
        }
    }

    console.log("DRAW");
    console.log(logSegments);
}

window.addEventListener("load", function() {

    AscFonts.g_fontApplication.Init();
    AscCommon.g_font_loader.put_Api(window.editor);

    AscFonts.load(null, function() {}, function() {});

    // FONTS
    let fontsContent = "";
    for (let i = 0; i < AscFonts.g_font_infos.length; i++)
    {
        fontsContent += ("<option value=\"" + AscFonts.g_font_infos[i].Name + "\">" + AscFonts.g_font_infos[i].Name + "</option>");
    }
    document.getElementById("id_fonts").innerHTML = fontsContent;

    // SCRIPTS
    let scriptContent = "";
    for (let item in AscFonts.HB_SCRIPT)
    {
        let text = item.substr("HB_SCRIPT_".length);
        text = text.substr(0, 1) + text.substr(1).toLowerCase();
        scriptContent += ("<option value=\"" + text + "\">" + text + "</option>");
    }
    document.getElementById("id_script").innerHTML = scriptContent;

    // DIRECTIONS
    let directionContent = "";
    directionContent += ("<option value=\"LTR\">LTR</option>");
    directionContent += ("<option value=\"RTL\">RTL</option>");
    directionContent += ("<option value=\"TTB\">TTB</option>");
    directionContent += ("<option value=\"BTT\">BTT</option>");
    document.getElementById("id_direction").innerHTML = directionContent;

    // AUTO
    document.getElementById("id_script_auto").onclick = function() {

        let script = getScriptByText();
        for (let item in AscFonts.HB_SCRIPT)
        {
            if (AscFonts.HB_SCRIPT[item] === script)
            {
                let text = item.substr("HB_SCRIPT_".length);
                text = text.substr(0, 1) + text.substr(1).toLowerCase();
                document.getElementById("id_script").value = text;
                break;
            }
        }
    };

    document.getElementById("id_direction_auto").onclick = function() {

        let script = getScriptByText();
        let direction = AscFonts.hb_get_script_horizontal_direction(script);

        if (direction === AscFonts.HB_DIRECTION.HB_DIRECTION_LTR)
            document.getElementById("id_direction").value = "LTR";
        else
            document.getElementById("id_direction").value = "RTL";
    };

    // DRAW
    document.getElementById("id_draw").onclick = function() {

        var codes = [];
        let text = document.getElementById("id_textarea").value;
        for (let iter = text.getUnicodeIterator(); iter.check(); iter.next())
        {
            codes.push(iter.value());
        }

        var currentFont = document.getElementById("id_fonts").value;
        var fonts = [new AscFonts.CFont(currentFont)];
        AscFonts.FontPickerByCharacter.checkTextLight(codes, true);
        AscFonts.FontPickerByCharacter.extendFonts(fonts);
        AscCommon.g_font_loader.LoadDocumentFonts2(fonts);
        return;
    };

});
