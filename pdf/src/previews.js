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
