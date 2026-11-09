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

(function (window) {
	// Same name Word writes on a header watermark, so a saved PDF can be
	// recognized, edited, and removed without treating the drawing as a
	// normal annotation.
	var WATERMARK_NAME = "PowerPlusWaterMarkObject";

	function isWatermarkName(sName) {
		return typeof sName === "string" && sName.indexOf(WATERMARK_NAME) === 0;
	}

	function watermarkOffset(nPageW, nPageH, nBoxW, nBoxH) {
		return {
			x: (nPageW - nBoxW) / 2,
			y: (nPageH - nBoxH) / 2
		};
	}

	// Word sizes an automatic watermark to the content box. A diagonal one
	// uses the page diagonal so the text stays inside a square page and a
	// wide page alike.
	function watermarkAutoWidth(nPageW, nPageH, nTextW, nTextH, bDiagonal) {
		if (!(nTextW > 0))
			return nPageW;
		if (bDiagonal)
			return Math.SQRT2 * Math.min(nPageW, nPageH) / (nTextH / nTextW + 1);
		return nPageW;
	}

	function watermarkRotation(bDiagonal) {
		return bDiagonal ? Math.PI * (360 - 45) / 180 : 0;
	}

	// The image dialog leaves opacity empty. Word still washes the picture,
	// and the semitransparent checkbox defaults to on, which is 128.
	function watermarkImageOpacity(nOpacity) {
		if (typeof nOpacity === "number" && nOpacity === nOpacity)
			return nOpacity;
		return 128;
	}

	// Paper stays white in the page bitmap, so an image drawn after it covers
	// the artwork. Keep the watermark only where the page is still paper.
	function compositeImageWatermarkBehind(aPage, aMark) {
		var nPaper = 250;
		var nCount = Math.min(aPage.length, aMark.length);
		for (var nIndex = 0; nIndex < nCount; nIndex += 4) {
			var nAlpha = aMark[nIndex + 3];
			if (!nAlpha)
				continue;
			if (aPage[nIndex] < nPaper || aPage[nIndex + 1] < nPaper || aPage[nIndex + 2] < nPaper)
				continue;
			var nCover = nAlpha / 255;
			var nKeep = 1 - nCover;
			aPage[nIndex] = (aMark[nIndex] * nCover + 255 * nKeep + 0.5) | 0;
			aPage[nIndex + 1] = (aMark[nIndex + 1] * nCover + 255 * nKeep + 0.5) | 0;
			aPage[nIndex + 2] = (aMark[nIndex + 2] * nCover + 255 * nKeep + 0.5) | 0;
			aPage[nIndex + 3] = 255;
		}
	}

	function isWatermarkDrawing(oDrawing) {
		if (!oDrawing || !oDrawing.getCNvProps)
			return false;
		var oPr = oDrawing.getCNvProps();
		return !!(oPr && isWatermarkName(oPr.name));
	}

	function nextWatermarkName() {
		if (window.AscFormat && window.AscFormat.Ax_Counter)
			return WATERMARK_NAME + (++window.AscFormat.Ax_Counter.GLOBAL_AX_ID_COUNTER);
		return WATERMARK_NAME;
	}

	function nameWatermark(oDrawing) {
		var oPr = oDrawing.getCNvProps && oDrawing.getCNvProps();
		if (oPr && oPr.setName)
			oPr.setName(nextWatermarkName());
	}

	function eachPageDrawing(oDoc, fVisit) {
		var nPages = oDoc.GetPagesCount();
		for (var i = 0; i < nPages; i++) {
			var oPage = oDoc.GetPageInfo(i);
			if (!oPage || !oPage.drawings)
				continue;
			var aDrawings = oPage.drawings.slice();
			for (var j = 0; j < aDrawings.length; j++)
				fVisit(aDrawings[j], i);
		}
	}

	function findPdfWatermark(oDoc) {
		var oFound = null;
		eachPageDrawing(oDoc, function (oDrawing) {
			if (!oFound && isWatermarkDrawing(oDrawing))
				oFound = oDrawing;
		});
		return oFound;
	}

	function removePdfWatermarks(oDoc) {
		eachPageDrawing(oDoc, function (oDrawing) {
			if (isWatermarkDrawing(oDrawing))
				oDoc.RemoveDrawing(oDrawing.GetId());
		});
	}

	// The writer stores fill opacity as a color alpha mod (value * 100000 / 255).
	function watermarkOpacityFromAlpha(nVal) {
		return Math.round(nVal * 255 / 100000);
	}

	function fillOpacity(oFill) {
		if (!oFill)
			return null;
		if (AscFormat.isRealNumber(oFill.transparent))
			return Math.round(oFill.transparent);
		var oColor = oFill.fill && oFill.fill.color;
		var aMods = oColor && oColor.Mods && oColor.Mods.Mods;
		if (!aMods)
			return null;
		for (var i = aMods.length - 1; i >= 0; i--) {
			if (aMods[i].name === "alpha" && AscFormat.isRealNumber(aMods[i].val))
				return watermarkOpacityFromAlpha(aMods[i].val);
		}
		return null;
	}

	function stripAlphaMod(oFill) {
		var oColor = oFill && oFill.fill && oFill.fill.color;
		var aMods = oColor && oColor.Mods && oColor.Mods.Mods;
		if (!aMods)
			return;
		for (var i = aMods.length - 1; i >= 0; i--) {
			if (aMods[i].name === "alpha")
				aMods.splice(i, 1);
		}
	}

	// The paragraph applicator keeps TextFill and deletes Unifill. The shape
	// writer saves Unifill only, so copy the fill onto the run itself.
	function persistWatermarkTextFill(oShape) {
		if (!isWatermarkDrawing(oShape) || !oShape.getDocContent)
			return;
		var oContent = oShape.getDocContent();
		if (!oContent || !oContent.Content)
			return;
		oContent.Content.forEach(function (oPara) {
			(oPara.Content || []).forEach(function (oRun) {
				if (!oRun.Pr || !oRun.Pr.TextFill || !oRun.Pr.TextFill.createDuplicate)
					return;
				oRun.Pr.Unifill = oRun.Pr.TextFill.createDuplicate();
			});
		});
	}

	// A reopened watermark has its opacity only as an alpha mod. Put it back on
	// TextFill, which is what the word-art painter reads, and keep auto-fit off.
	function restoreWatermarkTextFill(oShape) {
		if (!isWatermarkDrawing(oShape) || !oShape.getDocContent)
			return;
		var oContent = oShape.getDocContent();
		if (!oContent || !oContent.Content)
			return;
		var oSrc = null;
		var nOpacity = null;
		function remember(oFill) {
			if (oSrc || !oFill || !oFill.createDuplicate)
				return;
			var n = fillOpacity(oFill);
			if (!AscFormat.isRealNumber(n))
				return;
			oSrc = oFill;
			nOpacity = n;
		}
		oContent.Content.forEach(function (oPara) {
			(oPara.Content || []).forEach(function (oRun) {
				if (oRun.Pr)
					remember(oRun.Pr.Unifill || oRun.Pr.TextFill);
			});
			// A saved watermark keeps the fill on the paragraph mark, not on the letters.
			if (oPara.TextPr && oPara.TextPr.Value)
				remember(oPara.TextPr.Value.Unifill || oPara.TextPr.Value.TextFill);
		});
		if (!oSrc)
			return;
		oContent.Content.forEach(function (oPara) {
			(oPara.Content || []).forEach(function (oRun) {
				if (!oRun.Pr)
					return;
				var oFill = oSrc.createDuplicate();
				stripAlphaMod(oFill);
				oFill.transparent = nOpacity;
				oRun.Pr.TextFill = oFill;
				oRun.Pr.Unifill = oFill.createDuplicate();
			});
		});
		if (oShape.recalcInfo)
			oShape.recalcInfo.recalculateContent = true;
		if (oShape.getBodyPr && oShape.txBody && oShape.txBody.setBodyPr) {
			var oBodyPr = oShape.getBodyPr().createDuplicate();
			oBodyPr.textFit = new AscFormat.CTextFit();
			oBodyPr.textFit.type = AscFormat.text_fit_No;
			oShape.txBody.setBodyPr(oBodyPr);
		}
	}

	function applyWatermarkText(oShape, oProps) {
		var oContent = oShape.getDocContent();
		AscFormat.AddToContentFromString(oContent, oProps.get_Text() || "");

		var oTextPr = new CTextPr();
		var oTextPropMenu = oProps.get_TextPr();
		if (oTextPropMenu) {
			oTextPr.FontSize = (oTextPropMenu.get_FontSize() > 0 ? oTextPropMenu.get_FontSize() : 20);
			var oFamily = oTextPropMenu.get_FontFamily && oTextPropMenu.get_FontFamily();
			if (oFamily && oFamily.get_Name)
				oTextPr.RFonts.SetAll(oFamily.get_Name(), -1);
			oTextPr.Bold = oTextPropMenu.get_Bold();
			oTextPr.HighLight = oTextPropMenu.get_HighLight();
			oTextPr.Italic = oTextPropMenu.get_Italic();
			oTextPr.Underline = oTextPropMenu.get_Underline();
			oTextPr.Strikeout = oTextPropMenu.get_Strikeout();
			if (null !== oTextPropMenu.get_Lang())
				oTextPr.SetLang(oTextPropMenu.get_Lang());
		}
		// The shape record stores Unifill. TextFill is dropped on save, and the
		// paragraph applicator keeps only one of the two, so the opacity has to
		// be Unifill or the reopened letters come back with no fill at all.
		oTextPr.Unifill = AscFormat.CreateUnifillFromAscColor(oTextPropMenu && oTextPropMenu.get_Color() ? oTextPropMenu.get_Color() : new Asc.asc_CColor(0, 0, 0), 1);
		var nOpacity = oProps.get_Opacity();
		if (AscFormat.isRealNumber(nOpacity) && nOpacity < 255)
			oTextPr.Unifill.transparent = (nOpacity >> 0);

		oContent.SetApplyToAll(true);
		oContent.AddToParagraph(new ParaTextPr(oTextPr));
		oContent.SetParagraphAlign(AscCommon.align_Center);
		oContent.SetParagraphSpacing({Before: 0, After: 0, LineRule: Asc.linerule_Auto, Line: 1.0});
		oContent.SetParagraphIndent({FirstLine: 0, Left: 0, Right: 0});
		oContent.SetApplyToAll(false);

		var oBodyPr = oShape.getBodyPr().createDuplicate();
		oBodyPr.rot = 0;
		oBodyPr.spcFirstLastPara = false;
		oBodyPr.vertOverflow = AscFormat.nVOTOverflow;
		oBodyPr.horzOverflow = AscFormat.nHOTOverflow;
		oBodyPr.vert = AscFormat.nVertTThorz;
		oBodyPr.resetInsets();
		oBodyPr.numCol = 1;
		oBodyPr.spcCol = 0;
		oBodyPr.rtlCol = 0;
		oBodyPr.fromWordArt = false;
		oBodyPr.anchor = 1;
		oBodyPr.anchorCtr = false;
		oBodyPr.forceAA = false;
		oBodyPr.compatLnSpc = true;
		oBodyPr.prstTxWarp = null;
		// Auto-fit would rebuild the box from the run on the next recalc (Edit Text)
		// and the paragraph painter would then draw that run opaque.
		oBodyPr.textFit = new AscFormat.CTextFit();
		oBodyPr.textFit.type = AscFormat.text_fit_No;
		oBodyPr.wrap = AscFormat.nTWTNone;
		// PDF shapes read txBody.bodyPr. setBodyPr on the shape only stores the Word copy.
		if (oShape.txBody && oShape.txBody.setBodyPr)
			oShape.txBody.setBodyPr(oBodyPr);
		else
			oShape.setBodyPr(oBodyPr);
		return {content: oContent, textPr: oTextPr};
	}

	function addTextWatermark(oDoc, nPage, oProps) {
		var nPageW = oDoc.GetPageWidthMM(nPage);
		var nPageH = oDoc.GetPageHeightMM(nPage);
		var oTrack = new AscFormat.NewShapeTrack("textRect", 0, 0, oDoc.GetTheme(), undefined, undefined, oDoc, nPage);
		oTrack.track({}, 0, 0);
		var oShape = oTrack.getShape(false, oDoc.GetDrawingDocument(), oDoc.GetController());
		oShape.spPr.setFill(AscFormat.CreateNoFillUniFill());
		oShape.spPr.setLn(AscFormat.CreateNoFillLine());

		var oApplied = applyWatermarkText(oShape, oProps);
		var oContent = oApplied.content;
		var oTextPr = oApplied.textPr;
		var oXfrm = oShape.spPr.xfrm;
		var dRot = oProps.getXfrmRot();
		if (!AscFormat.fApproxEqual(0.0, dRot))
			oXfrm.setRot(dRot);

		var oContentSize = AscFormat.GetContentOneStringSizes(oContent);
		oContentSize.w = AscFormat.CTextBody.prototype.getContentWidth.call({content: oContent}) + 0.1;
		oXfrm.setExtX(oContentSize.w + 1);
		oXfrm.setExtY(oContentSize.h);

		var oTextPropMenu = oProps.get_TextPr();
		if (oTextPropMenu && oTextPropMenu.get_FontSize() < 0) {
			var nExtX = watermarkAutoWidth(nPageW, nPageH, oContentSize.w, oContentSize.h, !AscFormat.fApproxEqual(0.0, dRot));
			// The PDF text painter runs a little wider than the measure, and
			// the last letter wraps when the box is exact. Leave a margin.
			oTextPr.FontSize *= (nExtX / oContentSize.w) * 0.9;
			oContent.SetApplyToAll(true);
			oContent.AddToParagraph(new ParaTextPr(oTextPr));
			oContent.SetApplyToAll(false);
			oContentSize = AscFormat.GetContentOneStringSizes(oContent);
			oContentSize.w = AscFormat.CTextBody.prototype.getContentWidth.call({content: oContent}) + 0.1;
			oXfrm.setExtX(Math.max(nExtX, oContentSize.w + 2));
			oXfrm.setExtY(oContentSize.h);
		}

		var oPos = watermarkOffset(nPageW, nPageH, oXfrm.extX, oXfrm.extY);
		oXfrm.setOffX(oPos.x);
		oXfrm.setOffY(oPos.y);
		// Word paints a semitransparent watermark as word art. A PDF shape
		// otherwise draws the run at full opacity and the page shows through
		// only when this outline exists.
		if (oShape.recalcInfo)
			oShape.recalcInfo.recalculateContent = true;
		if (oShape.recalculateContent)
			oShape.recalculateContent();
		oDoc.AddDrawing(oShape, nPage, 0);
		nameWatermark(oShape);
		// Recalc keeps TextFill and drops Unifill. Put Unifill back on the run
		// so the shape record written at save still carries the fill.
		persistWatermarkTextFill(oShape);
	}

	function imageExtent(oDoc, oProps, nPageW, nPageH) {
		var oImgP = new Asc.asc_CImgProperty();
		oImgP.ImageUrl = oProps.get_ImageUrl();
		var oSize = oImgP.asc_getOriginSize(oDoc.GetApi());
		var dScale = oProps.get_Scale();
		var nImageW = oProps.get_ImageWidth();
		var nImageH = oProps.get_ImageHeight();
		var nExtX;
		var nExtY;
		if (AscFormat.isRealNumber(nImageW) && AscFormat.isRealNumber(nImageH) && nImageW > 0 && nImageH > 0) {
			nExtX = nImageW / 36000;
			nExtY = nImageH / 36000;
		} else if (dScale < 0) {
			nExtX = nPageW;
			nExtY = oSize.asc_getImageHeight() * (nExtX / oSize.asc_getImageWidth());
			if (nExtY > nPageH) {
				nExtY = nPageH;
				nExtX = oSize.asc_getImageWidth() * (nExtY / oSize.asc_getImageHeight());
			}
		} else {
			nExtX = oSize.asc_getImageWidth() * dScale;
			nExtY = oSize.asc_getImageHeight() * dScale;
		}
		return {x: nExtX, y: nExtY};
	}

	function addImageWatermark(oDoc, nPage, oProps) {
		var sUrl = oProps.get_ImageUrl();
		if (!sUrl)
			return;

		var nPageW = oDoc.GetPageWidthMM(nPage);
		var nPageH = oDoc.GetPageHeightMM(nPage);
		var oExt = imageExtent(oDoc, oProps, nPageW, nPageH);
		var oImage = new AscPDF.CPdfImage();
		AscFormat.fillImage(oImage, sUrl, 0, 0, oExt.x, oExt.y);

		var oXfrm = oImage.getXfrm();
		var dRot = oProps.getXfrmRot();
		if (!AscFormat.fApproxEqual(0.0, dRot))
			oXfrm.setRot(dRot);

		var nOpacity = watermarkImageOpacity(oProps.get_Opacity());
		if (nOpacity < 255 && oImage.blipFill)
			oImage.setTransparent(nOpacity);

		var oPos = watermarkOffset(nPageW, nPageH, oXfrm.extX, oXfrm.extY);
		oXfrm.setOffX(oPos.x);
		oXfrm.setOffY(oPos.y);
		oDoc.AddDrawing(oImage, nPage, 0);
		nameWatermark(oImage);
	}

	function addPdfWatermark(oDoc, nPage, oProps) {
		if (!oDoc.GetPageInfo(nPage))
			return;
		if (oProps.get_Type() === Asc.c_oAscWatermarkType.Image)
			addImageWatermark(oDoc, nPage, oProps);
		else if (oProps.get_Type() === Asc.c_oAscWatermarkType.Text)
			addTextWatermark(oDoc, nPage, oProps);
	}

	window.AscPDF = window.AscPDF || {};
	window.AscPDF.WATERMARK_NAME = WATERMARK_NAME;
	window.AscPDF.isWatermarkName = isWatermarkName;
	window.AscPDF.watermarkOffset = watermarkOffset;
	window.AscPDF.watermarkAutoWidth = watermarkAutoWidth;
	window.AscPDF.watermarkRotation = watermarkRotation;
	window.AscPDF.watermarkImageOpacity = watermarkImageOpacity;
	window.AscPDF.compositeImageWatermarkBehind = compositeImageWatermarkBehind;
	window.AscPDF.isWatermarkDrawing = isWatermarkDrawing;
	window.AscPDF.watermarkOpacityFromAlpha = watermarkOpacityFromAlpha;
	window.AscPDF.restoreWatermarkTextFill = restoreWatermarkTextFill;
	window.AscPDF.persistWatermarkTextFill = persistWatermarkTextFill;
	window.AscPDF.findPdfWatermark = findPdfWatermark;
	window.AscPDF.removePdfWatermarks = removePdfWatermarks;
	window.AscPDF.addPdfWatermark = addPdfWatermark;
})(window);
