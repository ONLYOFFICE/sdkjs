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
