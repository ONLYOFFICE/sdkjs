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
	 * Options to replace the HTML page content.
	 * @typedef {Object} ReplaceHtmlOptions
	 * @property {string} content - The HTML content to replace.
	 * @property {boolean} [separateParagraphs=true] - Specifies whether each paragraph will be created in a separate shape.
	 */

	/**
	 * Options to replace the XML page content.
	 * @typedef {Object} ReplaceXmlOptions
	 * @property {string[]} content - An array with XML shapes to be replaced.
	 */


    /**
     * Base class.
     * @global
     * @class
     * @name Api
     */
    let Api = window["PDFEditorApi"];

	/**
     * Returns the current page index.
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
	 * Returns the page image.
	 * @memberof Api
	 * @typeofeditors ["PDFE"]
	 * @alias GetPageImage
	 * @param {number} nPage - The page index.
	 * @param {object} [oParams={}] - The image parameters.
	 * @param {number} [oParams.maxSize] - The size of the larger image side in pixels.
	 * @param {boolean} [oParams.annotations=false] - Specifies whether the annotations will be rendered.
	 * @param {boolean} [oParams.fields=false] - Specifies whether the fields will be rendered.
	 * @param {boolean} [oParams.drawings=false] - Specifies whether the drawings will be rendered.
	 * @returns {canvas}
	 * @see office-js-api/Examples/Plugins/PDF/Api/Methods/GetPageImage.js
	 */
	Api.prototype["pluginMethod_GetPageImage"] = function(nPage, oParams) {
		if (!oParams) {
			oParams = {};
		}

		let oViewer		= this.DocumentRenderer;
		let oFile		= oViewer.file;
		let nPt2Px		= AscCommonWord.g_dKoef_pt_to_mm * AscCommonWord.g_dKoef_mm_to_pix;

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
     * Replaces the page content with the specified parameters.
     * @memberof Api
     * @typeofeditors ["PDFE"]
     * @alias ReplacePageContent
	 * @param {number} nPage - The page index.
	 * @param {object} oParams - The replacement parameters.
	 * @param {"xml" | "html"} oParams.type - The type of content to be replaced (XML / HTML).
	 * @param {ReplaceXmlOptions | ReplaceHtmlOptions} oParams.options - The content replacement options.
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
	/**
	 * Returns all the comments from the document.
	 * @memberof Api
	 * @typeofeditors ["PDFE"]
	 * @alias GetAllComments
	 * @returns {comment[]} - An array of comment objects containing the comment data.
	 * @see office-js-api/Examples/Plugins/PDF/Api/Methods/GetAllComments.js
	 */
	Api.prototype["pluginMethod_GetAllComments"] = function() {
		let oDoc = this.getPDFDoc();
		if (!oDoc)
			return;

		let arrResult = [];

		let aAnnots = oDoc.annots;
		for (let i = 0; i < aAnnots.length; i++) {
			let annot = aAnnots[i];

			let oAscCommData = annot.GetAscCommentData();
			let oCommData = oAscCommData ? new AscCommon.CCommentData() : undefined;
			oCommData && oCommData.Read_FromAscCommentData(oAscCommData);

			if (oCommData) {
				arrResult.push({"Id" : annot.GetId(), "Data" : oCommData.ConvertToSimpleObject()});
			}
		}

		return arrResult;
	};
})(window);


