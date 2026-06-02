/*
 * Copyright (C) Ascensio System SIA, 2009-2026
 *
 * This program is a free software product. You can redistribute it and/or
 * modify it under the terms of the GNU Affero General Public License (AGPL)
 * version 3 as published by the Free Software Foundation, together with the
 * additional terms provided in the LICENSE file.
 *
 * This program is distributed WITHOUT ANY WARRANTY; without even the implied
 * warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. For
 * details, see the GNU AGPL at: https://www.gnu.org/licenses/agpl-3.0.html
 *
 * You can contact Ascensio System SIA by email at info@onlyoffice.com
 * or by postal mail at 20A-6 Ernesta Birznieka-Upisha Street, Riga,
 * LV-1050, Latvia, European Union.
 *
 * The interactive user interfaces in modified versions of the Program
 * are required to display Appropriate Legal Notices in accordance with
 * Section 5 of the GNU AGPL version 3.
 *
 * No trademark rights are granted under this License.
 *
 * All non-code elements of the Product, including illustrations,
 * icon sets, and technical writing content, are licensed under the
 * Creative Commons Attribution-ShareAlike 4.0 International License:
 * https://creativecommons.org/licenses/by-sa/4.0/legalcode
 *
 * This license applies only to such non-code elements and does not
 * modify or replace the licensing terms applicable to the Program's
 * source code, which remains licensed under the GNU Affero General
 * Public License v3.
 *
 * SPDX-License-Identifier: AGPL-3.0-only
 */

"use strict";

(function(window, undefined){
    function CGraphics()
    {
        AscCommonExcel.CGraphics.call(this);
        this.m_oAutoShapesTrack = null;

        this.isDisableStrokeFillOptimization = true;
    }

    CGraphics.prototype = Object.create(AscCommonExcel.CGraphics.prototype);
    CGraphics.prototype.constructor = CGraphics;

    CGraphics.prototype.GetAutoShapesTrack = function()
    {
        return this.m_oAutoShapesTrack;
    };

    CGraphics.prototype.checkLoadingImage = function(img)
    {
        return (img.Status === AscFonts.ImageLoadStatus.Loading || (AscCommon.CollaborativeEditing.WaitImages && AscCommon.CollaborativeEditing.WaitImages[img]));
    };

    // TODO:
    CGraphics.prototype.font = function(font_id,font_size, matrix)
    {
        AscFonts.g_font_infos[AscFonts.g_map_font_index[font_id]].LoadFont(Asc.editor.FontLoader, this.m_oFontManager, font_size, 0, this.m_dDpiX, this.m_dDpiY, /*matrix*/undefined);
    };
    /*
    CGraphics.prototype.font = function(font_id,font_size)
    {
        AscFonts.g_font_infos[AscFonts.g_map_font_index[font_id]].LoadFont(Asc.editor.FontLoader, this.IsUseFonts2 ? this.m_oFontManager2 : this.m_oFontManager,
            Math.max(font_size, 1), 0, this.m_dDpiX, this.m_dDpiY, this.m_oTransform);
    };
    */

    CGraphics.prototype.SetTextPr = function(textPr, theme)
    {
        this.m_oTextPr = textPr.Copy();
        this.theme = theme;
        this.m_oTextPr.ReplaceThemeFonts(theme.themeElements.fontScheme);
    };

    CGraphics.prototype.GetRFonts = function()
    {
        return this.m_oTextPr.RFonts;
    };

    CGraphics.prototype.drawCollaborativeChanges = function(x, y, w, h)
    {
        this.b_color1( 0, 255, 0, 64 );
        this.rect( x, y, w, h );
        this.df();
    };

    CGraphics.prototype.drawSearchResult = function(x, y, w, h)
    {
        this.b_color1( 255, 220, 0, 200 );
        this.rect( x, y, w, h );
        this.df();
    };

    //------------------------------------------------------------export----------------------------------------------------
    window['AscCommon'] = window['AscCommon'] || {};
    window['AscCommon'].CGraphics = CGraphics;
})(window);
