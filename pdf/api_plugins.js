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

(function(window, undefined)
{
	/**
	 * Options for replace page content by html
	 * @typedef {Object} ReplaceHtmlOptions
	 * @property {string} content - html content to replace
	 * @property {boolean} [separateParagraphs=true] - will each paragraph be created a separate shape
	 */

	/**
	 * Options for replace page content by html
	 * @typedef {Object} ReplaceXmlOptions
	 * @property {string[]} content - array with xml shapes to replace
	 */


    /**
     * Base class.
     * @global
     * @class
     * @name Api
     */
    let Api = window["asc_docs_api"];

	/**
     * Gets current page index
     * @memberof Api
     * @typeofeditors ["PDFE"]
     * @alias GetCurrentPage
	 * @returns {number}
     * @see office-js-api/Examples/Plugins/PDF/Api/Methods/GetCurrentPage.js
	 */
    Api.prototype["pluginMethod_GetCurrentPage"] = function() {
        return this.getCurrentPage();
    };

	/**
	 * Gets page image
	 * @memberof Api
	 * @typeofeditors ["PDFE"]
	 * @alias GetPageImage
	 * @param {number} nPage - page index
	 * @param {object} [oParams={}] - image params
	 * @param {number} [oParams.maxSize] - the size of the larger side in pixels
	 * @param {boolean} [oParams.annotations=false] - will annotations be rendered
	 * @param {boolean} [oParams.fields=false] - will fields be rendered
	 * @param {boolean} [oParams.drawings=false] - will drawings be rendered
	 * @returns {canvas}
	 * @see office-js-api/Examples/Plugins/PDF/Api/Methods/GetPageImage.js
	 */
	Api.prototype["pluginMethod_GetPageImage"] = function(nPage, oParams) {
		if (!oParams) {
			oParams = {};
		}

		let oViewer		= this.DocumentRenderer;
		let oFile		= oViewer.file;
		let nPt2Px		= g_dKoef_pt_to_mm * g_dKoef_mm_to_pix;

		let nWidthPx	= oFile.pages[nPage].W * nPt2Px;
		let nHeightPx	= oFile.pages[nPage].H * nPt2Px;

		if (oParams["maxSize"]) {
			let aspectRatio = nWidthPx / nHeightPx;
			if (aspectRatio >= 1) {
				nWidthPx = oParams["maxSize"];
				nHeightPx = oParams["maxSize"] / aspectRatio;
			}
			else {
				nHeightPx = oParams["maxSize"];
				nWidthPx = oParams["maxSize"] * aspectRatio;
			}
		}

		nWidthPx = nWidthPx >> 0;
		nHeightPx = nHeightPx >> 0;

		let image = !oFile.pages[nPage].isRecognized ? oFile.getPage(nPage, nWidthPx, nHeightPx, undefined, 0xFFFFFF) : null;

		if (!image) {
			let pageColor = this.getPageBackgroundColor();

			image = document.createElement('canvas');

			let ctx = image.getContext('2d');

			image.width = nWidthPx;
			image.height = nHeightPx;

			ctx.fillStyle = "rgba(" + pageColor.R + "," + pageColor.G + "," + pageColor.B + ",1)";
			ctx.fillRect(0, 0, nWidthPx, nHeightPx);
		}

		image.requestWidth = nWidthPx;
		image.requestHeight = nHeightPx;

		let ctx = image.getContext('2d');

		if (oParams['drawings']) {
			oViewer._drawDrawingsOnCtx(nPage, ctx);
		}
		if (oParams['annotations']) {
			oViewer._drawMarkupAnnotsOnCtx(nPage, ctx);
			oViewer._drawAnnotsOnCtx(nPage, ctx);
		}
		if (oParams['fields']) {
			oViewer._drawFieldsOnCtx(nPage, ctx, false, true);
		}

		return ctx.canvas.toDataURL("image/png");
	};

    /**
     * Replace page content by params
     * @memberof Api
     * @typeofeditors ["PDFE"]
     * @alias ReplacePageContent
	 * @param {number} nPage - page index
	 * @param {object} oParams - replace params
	 * @param {"xml" | "html"} oParams.type - type of content to replace (xml / html)
	 * @param {ReplaceXmlOptions | ReplaceHtmlOptions} oParams.options - replace content options
	 * @returns {boolean}
     * @see office-js-api/Examples/Plugins/PDF/Api/Methods/ReplacePageContent.js
	 */
    Api.prototype["pluginMethod_ReplacePageContent"] = function(nPage, oParams) {
		let oDoc = this.getPDFDoc();
		
		let oPageInfo = oDoc.GetPageInfo(nPage);
		if (null == oPageInfo || (oPageInfo && (oPageInfo.IsDeleteLock() || oPageInfo.IsRecognized())) || true == this.isRestrictionView()) {
			return false;
		}

		if (oParams['type'] === 'xml') {
			return oDoc.EditPage(nPage, oParams['options']['content']);
		}
		else if (oParams['type'] === 'html') {
			if (oDoc.IsSelectionLocked(AscDFH.historydescription_Pdf_EditPage, [nPage])) {
				return false;
			}

			this.htmlPasteSepParagraphs = !!oParams["separateParagraphs"];
			this.htmlPastePageIdx = nPage;
			this['pluginMethod_PasteHtml'](oParams['options']['content']);
			oPageInfo.SetRecognized(true);
		}

		return true;
    };
})(window);


