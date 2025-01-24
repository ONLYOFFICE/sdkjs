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

(
    /**
     * @param {Window} window
     * @param {undefined} undefined
     */
    function(window, undefined) {

        function StampAnnotPreviewManager() {
        }

        StampAnnotPreviewManager.prototype.clear = function() {
            this.StampStyles.length = 0;
        };

        StampAnnotPreviewManager.prototype.getStampStyles = function() {
            if (this.StampStyles.length === 0) {
                this.generateStampStyles();
            }
            return this.StampStyles;
        };

        StampAnnotPreviewManager.prototype.createCanvas = function(width, height) {
            var oCanvas = document.createElement('canvas');
            oCanvas.width = AscCommon.AscBrowser.convertToRetinaValue(width, true);
            oCanvas.height = AscCommon.AscBrowser.convertToRetinaValue(height, true);
            return oCanvas;
        };

        StampAnnotPreviewManager.prototype.getStampRender = function(sType) {
			let oDoc = Asc.editor.getPDFDoc();
            let oTextDrawer = oDoc.CreateStampRender(sType);

			return oTextDrawer;
        };

        StampAnnotPreviewManager.prototype.getStampPreview = function(type) {
            return this.getStampPreviewCanvas(type);
        };
        StampAnnotPreviewManager.prototype.getStampPreviews = function() {
			let aRet = [];

			let _t = this;
            for (let type in AscPDF.STAMPS_JSON) {
                let oPreview = {};
                oPreview["Type"] = type;
                oPreview["Image"] = _t.getStampPreview(type);
                aRet.push(oPreview);
            }

			return aRet;
        };
        StampAnnotPreviewManager.prototype.getStampPreviewCanvas = function(type) {
            let oTextDrawer	= this.getStampRender(type);
            if (!oTextDrawer) {
                return "";
            }

			let pxWidth = oTextDrawer.Width * g_dKoef_mm_to_pix;
            let pxHeight = oTextDrawer.Height * g_dKoef_mm_to_pix;

			let canvas		= this.createCanvas(pxWidth / 2 >> 0, pxHeight / 2 >> 0);
            let ctx			= canvas.getContext('2d');
            let graphics	= new AscCommon.CGraphics();
            
            graphics.init(ctx, canvas.width, canvas.height, oTextDrawer.Width, oTextDrawer.Height);
            graphics.m_oFontManager = AscCommon.g_fontManager;
            graphics.transform(1, 0, 0, 1, 0, 0);

            oTextDrawer.m_aStack[0].draw(graphics);
            return canvas;
        };

        //----------------------------------------------------------export----------------------------------------------------
        window['AscPDF'] = window['AscPDF'] || {};
        window['AscPDF'].StampAnnotPreviewManager = StampAnnotPreviewManager;
    })(window);
